import pypsa
import xarray as xr
import numpy as np
import pandas as pd


def _haversine_km(lon0, lat0, lon1, lat1):
    """
    Distancia haversine entre dos puntos en km.
    lon/lat en grados.
    """
    R = 6371.0  # km

    lon0 = np.radians(float(lon0))
    lat0 = np.radians(float(lat0))
    lon1 = np.radians(float(lon1))
    lat1 = np.radians(float(lat1))

    dlon = lon1 - lon0
    dlat = lat1 - lat0

    a = (
        np.sin(dlat / 2.0) ** 2
        + np.cos(lat0) * np.cos(lat1) * np.sin(dlon / 2.0) ** 2
    )

    c = 2 * np.arcsin(np.sqrt(a))

    return R * c


def _is_physical_ac_line(line_name, grid: pypsa.Network) -> bool:
    """
    Identifica líneas añadidas por add_lines():

        name = L{desde}_{hasta}
        bus0 = Bus.{desde}
        bus1 = Bus.{hasta}

    Excluye líneas/buses auxiliares.
    """
    if line_name not in grid.lines.index:
        return False

    if not str(line_name).startswith("L"):
        return False

    bus0 = str(grid.lines.at[line_name, "bus0"])
    bus1 = str(grid.lines.at[line_name, "bus1"])

    if not bus0.startswith("Bus."):
        return False

    if not bus1.startswith("Bus."):
        return False

    if bus0 not in grid.buses.index:
        return False

    if bus1 not in grid.buses.index:
        return False

    return True


def _calculate_physical_line_lengths_km(
    grid: pypsa.Network,
    physical_line_names: pd.Index,
) -> pd.Series:
    """
    Calcula la distancia de las líneas físicas usando las coordenadas
    de los buses conectados.
    """
    lengths = {}

    for line_name in physical_line_names:
        bus0 = grid.lines.at[line_name, "bus0"]
        bus1 = grid.lines.at[line_name, "bus1"]

        lon0 = grid.buses.at[bus0, "x"]
        lat0 = grid.buses.at[bus0, "y"]
        lon1 = grid.buses.at[bus1, "x"]
        lat1 = grid.buses.at[bus1, "y"]

        if pd.isna(lon0) or pd.isna(lat0) or pd.isna(lon1) or pd.isna(lat1):
            lengths[line_name] = np.nan
        else:
            lengths[line_name] = _haversine_km(lon0, lat0, lon1, lat1)

    return pd.Series(lengths, dtype=float)


def add_line_flow_penalty(
    grid: pypsa.Network,
    snapshots,
    penalty_eur_per_mwh: float = 0.1,
    use_length_scaling: bool = False,
) -> None:
    """
    Añade una penalización lineal al valor absoluto del flujo por las líneas físicas AC.

    Solo se penalizan las líneas creadas por add_lines(), es decir:
        - líneas cuyo nombre empieza por "L"
        - conectadas entre buses cuyo nombre empieza por "Bus."

    Coste añadido:
        sum_t,l weight_t * penalty_l * |flow_l,t|

    No representa pérdidas físicas reales, sino un coste proxy por uso de red.
    """

    m = grid.model

    if "Line-s" not in m.variables:
        print("WARNING: No se encuentra la variable 'Line-s'. No se añade penalización de líneas.")
        print("Variables disponibles:", list(m.variables))
        return

    line_flow_all = m.variables["Line-s"]

    # Detectar dimensión de líneas en Linopy/PyPSA.
    non_snapshot_dims = [d for d in line_flow_all.dims if d != "snapshot"]

    if len(non_snapshot_dims) != 1:
        raise ValueError(
            f"No se ha podido identificar la dimensión de líneas en Line-s. "
            f"Dims encontradas: {line_flow_all.dims}"
        )

    line_dim = non_snapshot_dims[0]
    all_line_names = line_flow_all.coords[line_dim].to_index()

    # Filtrar solo líneas físicas creadas por add_lines()
    physical_line_names = pd.Index([
        line_name
        for line_name in all_line_names
        if _is_physical_ac_line(line_name, grid)
    ])

    if physical_line_names.empty:
        print("WARNING: No se han encontrado líneas físicas AC para penalizar.")
        return

    # Seleccionar solo flujos de líneas físicas
    line_flow = line_flow_all.sel({line_dim: physical_line_names})

    # Variable auxiliar: abs_line_flow >= |line_flow|
    abs_line_flow = m.add_variables(
        lower=0,
        coords=line_flow.coords,
        name="Physical-Line-abs-flow",
    )

    m.add_constraints(
        abs_line_flow >= line_flow,
        name="physical_line_abs_flow_positive",
    )

    m.add_constraints(
        abs_line_flow >= -line_flow,
        name="physical_line_abs_flow_negative",
    )

    # Pesos temporales en horas
    weights = xr.DataArray(
        grid.snapshot_weightings.objective.loc[snapshots].to_numpy(),
        dims=["snapshot"],
        coords={"snapshot": snapshots},
    )

    # Penalización por línea
    if use_length_scaling:
        line_lengths_km = _calculate_physical_line_lengths_km(
            grid,
            physical_line_names,
        )

        positive_lengths = line_lengths_km[line_lengths_km > 0]

        if positive_lengths.empty:
            print(
                "WARNING: No hay longitudes válidas para las líneas físicas. "
                "Se aplica penalización uniforme."
            )
            line_penalty_values = pd.Series(
                penalty_eur_per_mwh,
                index=physical_line_names,
                dtype=float,
            )
        else:
            fallback_length = positive_lengths.mean()

            line_lengths_km = line_lengths_km.fillna(fallback_length)
            line_lengths_km = line_lengths_km.where(
                line_lengths_km > 0,
                fallback_length,
            )

            mean_length = line_lengths_km.mean()

            line_penalty_values = (
                penalty_eur_per_mwh * line_lengths_km / mean_length
            )

        line_penalty = xr.DataArray(
            line_penalty_values.to_numpy(),
            dims=[line_dim],
            coords={line_dim: physical_line_names.to_numpy()},
        )

        print("\n=== Physical line length penalty info ===")
        print(line_lengths_km.describe())
        print("Mean line penalty:", line_penalty_values.mean())
        print("Min line penalty:", line_penalty_values.min())
        print("Max line penalty:", line_penalty_values.max())

    else:
        line_penalty = penalty_eur_per_mwh

    penalty_term = (abs_line_flow * weights * line_penalty).sum()

    m.objective += penalty_term

    print(
        f"Physical line flow penalty added: {penalty_eur_per_mwh} €/MWh, "
        f"use_length_scaling={use_length_scaling}, "
        f"n_lines={len(physical_line_names)}, "
        f"line_dim={line_dim}"
    )