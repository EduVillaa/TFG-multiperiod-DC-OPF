from __future__ import annotations

import logging
import warnings

if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()

import matplotlib
matplotlib.use("Agg")

import pandas as pd
import pypsa
from pathlib import Path
import sys
from datetime import datetime, timedelta
import threading
import time
import traceback
from types import SimpleNamespace


from Postprocessing.export_multiperiod_results import export_multiperiod_results, extract_multiperiod_result_tables
from Postprocessing.export_static_results import export_static_results
from Network_builder.Network.build_network import build_network
from Network_builder.Network.buses import add_buses
from Network_builder.Network.grid_connection import *
from Network_builder.Network.lines import add_lines
from Network_builder.Network.loads import *
from Network_builder.Generators.PVandWindGenerators import *
from Network_builder.Generators.dispatchable import add_dispatchable_generators
from Network_builder.Storage.constraints import add_battery_constraints
from Network_builder.Storage.storage_model import add_storage_as_store_links
from Postprocessing.drawgridinmap import drawrealgrid
from Network_builder.Network.Load_Profiles_SPAIN import build_hourly_demand_by_region, build_monthly_nodal_load_weights_ES, build_hourly_nodal_demand, node_to_region
from Network_builder.Network.Load_Profiles_PT import regional_hourly_demand_builder, build_monthly_nodal_load_weights_PT, node_to_region_PT, build_hourly_nodal_demand_PT
from Network_builder.Generators.PVandWind_profiles import renewable_profile_builder
from Network_builder.Generators.GasPriceBuilder import CCGT_dataframe_treatment, daily_to_snapshots
from Network_builder.Storage.runoff4hydro import build_hydro_inflow
from Network_builder.Storage.llenado_embalses4hydro import get_embalses_closest_date
from Network_builder.Generators.runoff4ror import build_ror_p_max_pu
from Network_builder.Network.lineflowpenalty import add_line_flow_penalty

def leerhojas(filename: str | Path) -> dict:
    sheets = {}
    
    sheets["Net_Buses"] = pd.read_excel(
        filename,
        sheet_name="Net_Buses",
        header=1
    ).iloc[:, 1:]

    sheets["Net_Lines"] = pd.read_excel(
        filename,
        sheet_name="Net_Lines",
        header=2
    ).iloc[:, 1:]

    sheets["Gen_Dispatchable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Dispatchable",
        header=2
    ).iloc[:, 1:]

    sheets["Gen_PV_and_Wind"] = pd.read_excel(
        filename,
        sheet_name="Gen_PV_and_Wind",
        header=2
    ).iloc[:, 1:]

    sheets["StorageUnit"] = pd.read_excel(
        filename,
        sheet_name="StorageUnit",
        header=2
    ).iloc[:, 1:]

    sheets["Grid_connection"] = pd.read_excel(
        filename,
        sheet_name="Grid_connection",
        header=2
    ).iloc[:, 1:]

    return sheets

def build_sys_settings_from_gui(gui_params: dict) -> pd.DataFrame:
    # Parámetros siempre obligatorios
    base_required = [
        "VOLL (€/MWh)",
        "Static / Multiperiod",
        "line_flow_penalty",
        "use_line_length_scaling",
        "solver_name",
        "mip_rel_gap",
        "time_limit",
        "Notes",
    ]

    # Validar básicos
    missing_base = [k for k in base_required if k not in gui_params]
    if missing_base:
        raise KeyError(f"Faltan parámetros básicos de la GUI: {missing_base}")

    horizon = gui_params["Static / Multiperiod"]
    solver_name = str(gui_params.get("solver_name", "highs")).lower()

    if solver_name not in ["highs", "gurobi"]:
        raise ValueError(
            f"Solver no válido: {solver_name}. "
            "Debe ser 'highs' o 'gurobi'."
        )

    data = {
        "VOLL (€/MWh)": float(gui_params["VOLL (€/MWh)"]),
        "Static / Multiperiod": horizon,
        "solver_name": solver_name,
        "mip_rel_gap": (
            None
            if gui_params.get("mip_rel_gap") is None
            else float(gui_params.get("mip_rel_gap"))
        ),
        "time_limit": int(gui_params.get("time_limit", 3600)),
        "threads": int(gui_params.get("threads", 0)),

        "mip_focus": (
            None
            if gui_params.get("mip_focus") is None
            else int(gui_params.get("mip_focus"))
        ),

        "method": (
            None
            if gui_params.get("method") is None
            else int(gui_params.get("method"))
        ),
        "crossover": (
            None
            if gui_params.get("crossover") is None
            else int(gui_params.get("crossover"))
        ),
        "numeric_focus": (
        None
        if gui_params.get("numeric_focus") is None
        else int(gui_params.get("numeric_focus"))
        ),


        "bar_conv_tol": (
            None
            if gui_params.get("bar_conv_tol") is None
            else float(gui_params.get("bar_conv_tol"))
        ),
        "bar_homogeneous": (
            None
            if gui_params.get("bar_homogeneous") is None
            else int(gui_params.get("bar_homogeneous"))
        ),
        "feasibility_tol": (
            None
            if gui_params.get("feasibility_tol") is None
            else float(gui_params.get("feasibility_tol"))
        ),
        "optimality_tol": (
            None
            if gui_params.get("optimality_tol") is None
            else float(gui_params.get("optimality_tol"))
        ),

        "line_flow_penalty": float(gui_params.get("line_flow_penalty", 0.0)),
        "use_line_length_scaling": bool(gui_params.get("use_line_length_scaling", True)),
        "Notes": str(gui_params.get("Notes", "")),
    }

    if horizon == "Multiperiod":
        multiperiod_required = [
            "Start date (dd/mm/aaaa)",
            "Simulation duration (days)",
            "Graph resolution",
        ]

        missing_multi = [k for k in multiperiod_required if k not in gui_params]
        if missing_multi:
            raise KeyError(f"Faltan parámetros de Multiperiod: {missing_multi}")

        data.update({
            "Start date (dd/mm/aaaa)": pd.to_datetime(gui_params["Start date (dd/mm/aaaa)"]),
            "Simulation duration (days)": int(gui_params["Simulation duration (days)"]),
            "Graph resolution": gui_params["Graph resolution"],

            # No aplica en multiperiodo
            "Static snapshot datetime": None,
        })

        intermediate_enabled_raw = gui_params.get(
            "intermediate_storage_constraints_enabled",
            False,
        )
        intermediate_enabled = (
            str(intermediate_enabled_raw).strip().lower()
            in ["true", "1", "yes", "sí", "si"]
        )
        if intermediate_enabled:
            data.update({
                "intermediate_storage_constraints_enabled": True,
                "intermediate_storage_constraint_days": int(
                    gui_params.get("intermediate_storage_constraint_days", 3)
                ),
                "intermediate_hydro_soc_band_percent": float(
                    gui_params.get("intermediate_hydro_soc_band_percent", 0.5)
                ),
            })

    elif horizon == "Static":
        static_required = [
            "Static snapshot datetime",
        ]

        missing_static = [k for k in static_required if k not in gui_params]
        if missing_static:
            raise KeyError(f"Faltan parámetros de Static: {missing_static}")

        data.update({
            # No aplica en estático
            "Start date (dd/mm/aaaa)": None,
            "Simulation duration (days)": None,
            "Graph resolution": None,

            # Snapshot concreto para OPF estático
            "Static snapshot datetime": pd.to_datetime(gui_params["Static snapshot datetime"]),
        })

    else:
        raise ValueError(
            f"Valor no válido para 'Static / Multiperiod': {horizon}. "
            "Debe ser 'Static' o 'Multiperiod'."
        )

    df_SYS_settings = pd.DataFrame({
        "SYSTEM PARAMETERS": pd.Series(data)
    })

    return df_SYS_settings

def build_battery_economic_settings_from_gui(gui_params: dict) -> pd.DataFrame:
    # Parámetro base imprescindible para decidir si aplica o no
    if "Static / Multiperiod" not in gui_params:
        raise KeyError("Falta el parámetro básico de la GUI: 'Static / Multiperiod'")

    horizon = gui_params["Static / Multiperiod"]

    # Detectar si los parámetros económicos de batería están presentes
    has_discount_rate = "Discount rate (%)" in gui_params
    has_battery_lifetime = "Default battery lifetime (years)" in gui_params

    # Solo exigimos ambos si estamos en multiperiod y aparece al menos uno de los dos
    # o si el flujo de la GUI los ha incluido explícitamente
    if horizon == "Multiperiod" and (has_discount_rate or has_battery_lifetime):
        missing = []
        if not has_discount_rate:
            missing.append("Discount rate (%)")
        if not has_battery_lifetime:
            missing.append("Default battery lifetime (years)")

        if missing:
            raise KeyError(f"Faltan parámetros económicos de batería: {missing}")

        data = {
            "Discount rate (%)": float(gui_params["Discount rate (%)"]),
            "Default battery lifetime (years)": int(gui_params["Default battery lifetime (years)"]),
        }

    else:
        # Caso Static o Multiperiod sin optimización de batería
        data = {
            "Discount rate (%)": None,
            "Default battery lifetime (years)": None,
        }

    df_battery_economic_settings = pd.DataFrame({
        "BATTERY ECONOMIC PARAMETERS": pd.Series(data)
    })

    return df_battery_economic_settings


def start_heartbeat(message="Optimización en curso", interval=900):
    """
    Imprime un mensaje periódico mientras una tarea larga sigue ejecutándose para que no parezca que ha habido un error
    para interval 900 son 15 minutos
    """
    stop_event = threading.Event()

    def heartbeat():
        start = time.time()

        while not stop_event.wait(interval):
            elapsed_min = (time.time() - start) / 60
            print(
                f"[heartbeat] {message}... "
                f"{elapsed_min:.1f} min transcurridos",
                flush=True,
            )

    thread = threading.Thread(target=heartbeat, daemon=True)
    thread.start()

    return stop_event

def solve_opf(
    grid,
    solver_name,
    battery_specs=None,
    final_hydro_soc_fraction=None,
    initial_hydro_soc_fraction=None,
    mip_rel_gap=None,
    time_limit=None,
    threads=None,
    mip_focus=None,
    method=None,
    crossover=None,
    numeric_focus=None,
    bar_conv_tol=None,
    feasibility_tol=None,
    optimality_tol=None,
    bar_homogeneous=None,
    line_flow_penalty: float = 0.0,
    use_line_length_scaling: bool = True,
    additional_extra_functionality=None,
) -> tuple:
    """
    Resuelve el OPF con hasta tres intentos para la restricción de SOC final hidro:

    1) SOC_final >= SOC_final_objetivo
    2) Si es infactible: SOC_final >= SOC_inicial
    3) Si también es infactible: sin restricción de SOC final hidro

    Returns
    -------
    status : str
        Estado devuelto por PyPSA/Linopy.
    condition : str
        Condición de la optimización: optimal, infeasible, etc.
    hydro_soc_mode : str
        Modo finalmente usado:
        - "target_soc"
        - "initial_soc"
        - "no_hydro_final_soc"
        - "failed"
    """

    solver_name = str(solver_name).lower()

    if solver_name not in ["highs", "gurobi"]:
        raise ValueError(
            f"Solver no soportado: {solver_name}. "
            "Debe ser 'highs' o 'gurobi'."
        )

    solver_options = {}

    # ---------------------------------------------------------
    # Opciones específicas de cada solver
    # ---------------------------------------------------------
    if solver_name == "highs":
        if mip_rel_gap is not None:
            solver_options["mip_rel_gap"] = mip_rel_gap

        if time_limit is not None:
            solver_options["time_limit"] = time_limit

        # Opciones extra de log para HiGHS
        solver_options["output_flag"] = True
        solver_options["log_to_console"] = True
        solver_options["mip_min_logging_interval"] = 30

    elif solver_name == "gurobi":
        if mip_rel_gap is not None:
            solver_options["MIPGap"] = mip_rel_gap

        if time_limit is not None:
            solver_options["TimeLimit"] = time_limit

        if threads is not None:
            solver_options["Threads"] = threads

        if mip_focus is not None:
            solver_options["MIPFocus"] = mip_focus
        
        if method is not None:
            solver_options["Method"] = method
        
        if bar_homogeneous is not None:
            solver_options["BarHomogeneous"] = bar_homogeneous
        
        if crossover is not None:
            solver_options["Crossover"] = crossover

        if numeric_focus is not None:
            solver_options["NumericFocus"] = numeric_focus

        if bar_conv_tol is not None:
            solver_options["BarConvTol"] = bar_conv_tol

        if feasibility_tol is not None:
            solver_options["FeasibilityTol"] = feasibility_tol

        if optimality_tol is not None:
            solver_options["OptimalityTol"] = optimality_tol

        # Opcional: deja log en consola
        solver_options["OutputFlag"] = 1

    print(f"Resolviendo OPF con solver: {solver_name}")
    print(f"Opciones del solver: {solver_options}")

    def make_extra_functionality(hydro_soc_fraction, attempt_name):
        def extra_functionality(n, snapshots):

            # Restricciones de batería
            if battery_specs is not None:
                add_battery_constraints(n, snapshots, battery_specs)

            # Restricción agregada de SOC final hidro
            if hydro_soc_fraction is not None:
                add_hydro_final_soc_constraint(
                    n,
                    snapshots,
                    final_soc_fraction=hydro_soc_fraction,
                    constraint_name=f"hydro_final_soc_min_{attempt_name}",
                )

            if additional_extra_functionality is not None:
                additional_extra_functionality(n, snapshots)

            # Penalización al flujo en las líneas
            if line_flow_penalty is not None and line_flow_penalty > 0:
                add_line_flow_penalty(
                    n,
                    snapshots,
                    penalty_eur_per_mwh=line_flow_penalty,
                    use_length_scaling=use_line_length_scaling,
                )

        return extra_functionality

    def is_success(status, condition):
        return status == "ok" and str(condition).lower() in ["optimal", "suboptimal"]

    def is_infeasible(condition):
        return str(condition).lower() in ["infeasible", "infeasible_or_unbounded"]

    def run_optimization(extra_functionality, attempt_label):
        """
        Ejecuta grid.optimize() con heartbeat para que la terminal muestre
        señales de vida aunque el solver esté mucho tiempo sin escribir.
        """
        heartbeat_stop = start_heartbeat(
            message=f"{solver_name} sigue resolviendo ({attempt_label})",
            interval=900,  # cada 15 minutos
        )

        try:
            status, condition = grid.optimize(
                solver_name=solver_name,
                extra_functionality=extra_functionality,
                solver_options=solver_options,
            )
        finally:
            heartbeat_stop.set()

        return status, condition

    # ---------------------------------------------------------
    # Intento 1: SOC_final >= SOC_final_objetivo
    # ---------------------------------------------------------
    if final_hydro_soc_fraction is not None:
        print(
            "Intento 1: resolviendo con "
            f"SOC_final >= SOC_final_objetivo = {final_hydro_soc_fraction:.4f}",
            flush=True,
        )

        status, condition = run_optimization(
            extra_functionality=make_extra_functionality(
                hydro_soc_fraction=final_hydro_soc_fraction,
                attempt_name="target",
            ),
            attempt_label="intento 1: SOC final objetivo",
        )

        if is_success(status, condition):
            print("Optimización exitosa con restricción SOC final objetivo.")
            return status, condition, "target_soc"

        if not is_infeasible(condition):
            print(
                "Optimización fallida por una causa distinta a infactibilidad: "
                f"status={status}, condition={condition}"
            )
            return status, condition, "failed"

        print(
            "El intento 1 ha resultado infactible con "
            "SOC_final >= SOC_final_objetivo."
        )

    else:
        print(
            "No se ha proporcionado final_hydro_soc_fraction. "
            "Se omite el intento 1."
        )

    # ---------------------------------------------------------
    # Intento 2: SOC_final >= SOC_inicial
    # ---------------------------------------------------------
    if initial_hydro_soc_fraction is not None:
        print(
            "Intento 2: reintentando con "
            f"SOC_final >= SOC_inicial = {initial_hydro_soc_fraction:.4f}",
            flush=True,
        )

        status, condition = run_optimization(
            extra_functionality=make_extra_functionality(
                hydro_soc_fraction=initial_hydro_soc_fraction,
                attempt_name="initial",
            ),
            attempt_label="intento 2: SOC final inicial",
        )

        if is_success(status, condition):
            print("Optimización exitosa con restricción SOC_final >= SOC_inicial.")
            return status, condition, "initial_soc"

        if not is_infeasible(condition):
            print(
                "Optimización fallida por una causa distinta a infactibilidad: "
                f"status={status}, condition={condition}"
            )
            return status, condition, "failed"

        print(
            "El intento 2 también ha resultado infactible con "
            "SOC_final >= SOC_inicial."
        )

    else:
        print(
            "No se ha proporcionado initial_hydro_soc_fraction. "
            "Se omite el intento 2."
        )

    # ---------------------------------------------------------
    # Intento 3: sin restricción de SOC final hidro
    # ---------------------------------------------------------
    if additional_extra_functionality is not None:
        print(
            "Intento 3: resolviendo sin restricción hidro normal; "
            "se aplican restricciones extra proporcionadas...",
            flush=True,
        )
    else:
        print("Intento 3: reintentando sin restricción de SOC final hidro...", flush=True)

    status, condition = run_optimization(
        extra_functionality=make_extra_functionality(
            hydro_soc_fraction=None,
            attempt_name="none",
        ),
        attempt_label="intento 3: sin SOC final hidro",
    )

    if is_success(status, condition):
        if additional_extra_functionality is not None:
            print("Optimización exitosa con restricciones extra proporcionadas.")
        else:
            print("Optimización exitosa sin restricción de SOC final hidro.")
        return status, condition, "no_hydro_final_soc"

    print(
        "Optimización fallida también sin restricción de SOC final hidro: "
        f"status={status}, condition={condition}"
    )

    return status, condition, "failed"


def add_hydro_final_soc_constraint(
    grid,
    snapshots,
    final_soc_fraction,
    constraint_name="hydro_final_soc_min",
):
    m = grid.model

    hydro_stores = grid.stores.index[grid.stores.carrier == "hydro"]

    if len(hydro_stores) == 0:
        print("No hay stores hidroeléctricos. No se añade restricción de SOC final.")
        return

    final_snapshot = snapshots[-1]

    hydro_soc_final = m.variables["Store-e"].loc[final_snapshot, hydro_stores]

    hydro_e_nom_total = grid.stores.loc[hydro_stores, "e_nom"].sum()

    m.add_constraints(
        hydro_soc_final.sum() >= final_soc_fraction * hydro_e_nom_total,
        name=constraint_name,
    )


def _constraint_name_suffix(value) -> str:
    return "".join(char if char.isalnum() else "_" for char in str(value))


def _embalses_soc_fraction(df_embalses_soc: pd.DataFrame, label: str) -> float:
    required_columns = {"AGUA_ACTUAL", "AGUA_TOTAL"}
    missing_columns = required_columns - set(df_embalses_soc.columns)
    if missing_columns:
        raise ValueError(
            f"Faltan columnas de embalses para calcular SOC {label}: "
            f"{sorted(missing_columns)}"
        )

    if df_embalses_soc.empty:
        raise ValueError(f"No hay datos de embalses para calcular SOC {label}.")

    agua_actual = pd.to_numeric(
        df_embalses_soc["AGUA_ACTUAL"],
        errors="coerce",
    ).sum()
    agua_total = pd.to_numeric(
        df_embalses_soc["AGUA_TOTAL"],
        errors="coerce",
    ).sum()

    if pd.isna(agua_total) or agua_total <= 0:
        raise ValueError(
            f"AGUA_TOTAL no válido para calcular SOC {label}: {agua_total}"
        )

    fraction = float(agua_actual / agua_total)
    if not 0.0 <= fraction <= 1.0:
        raise ValueError(
            f"SOC {label} fuera de rango [0, 1]: {fraction:.6f}"
        )

    return fraction


def _infer_rolling_window_step(rolling_windows) -> pd.Timedelta:
    for window in rolling_windows:
        window = pd.DatetimeIndex(window)
        if len(window) > 1:
            step = pd.Timestamp(window[1]) - pd.Timestamp(window[0])
            if step <= pd.Timedelta(0):
                raise ValueError("Los snapshots rolling horizon no están ordenados.")
            return step

    return pd.Timedelta(hours=1)


def build_hydro_soc_target_trajectory(
    startdate,
    fecha_final,
    rolling_windows,
    df_embalses_original: pd.DataFrame,
    hydro_band_fraction: float,
    log_prefix: str = "Rolling horizon",
) -> dict:
    # Rolling horizon: trayectoria terminal agregada para stores con carrier hydro.
    hydro_band_fraction = float(hydro_band_fraction)
    if hydro_band_fraction < 0.0 or hydro_band_fraction > 0.30:
        raise ValueError("rolling_hydro_soc_band debe estar entre 0 y 0.30.")

    t0 = pd.to_datetime(startdate)
    t1 = pd.to_datetime(fecha_final)

    if pd.isna(t0) or pd.isna(t1):
        raise ValueError("Fechas no válidas para construir trayectoria hydro rolling horizon.")

    if t1 <= t0:
        raise ValueError(
            "El horizonte rolling horizon debe tener fecha final posterior a la inicial."
        )

    rolling_windows = [pd.DatetimeIndex(window) for window in rolling_windows]
    if not rolling_windows:
        raise ValueError("No hay ventanas para construir la trayectoria hydro.")

    df_embalses_initial_SOC = get_embalses_closest_date(
        df=df_embalses_original,
        target_date=t0,
    )
    df_embalses_final_SOC = get_embalses_closest_date(
        df=df_embalses_original,
        target_date=t1,
    )

    target_start = _embalses_soc_fraction(df_embalses_initial_SOC, "inicial")
    target_end = _embalses_soc_fraction(df_embalses_final_SOC, "final")

    print(
        f"{log_prefix} hydro SOC inicial calculado: "
        f"{target_start:.4f} ({target_start:.2%})",
        flush=True,
    )
    print(
        f"{log_prefix} hydro SOC final objetivo calculado: "
        f"{target_end:.4f} ({target_end:.2%})",
        flush=True,
    )
    print(
        f"Margen trayectoria hydro {log_prefix} elegido: "
        f"{hydro_band_fraction:.4f} ({hydro_band_fraction:.2%})",
        flush=True,
    )

    snapshot_step = _infer_rolling_window_step(rolling_windows)
    total_duration = t1 - t0
    trajectory = {}

    for window_number, window in enumerate(rolling_windows, start=1):
        if window.empty:
            raise ValueError("Se ha recibido una ventana rolling horizon vacía.")

        window_label = f"rolling_window_{window_number:03d}"
        constraint_snapshot = pd.Timestamp(window[-1])
        window_end = min(constraint_snapshot + snapshot_step, t1)

        alpha = float((window_end - t0) / total_duration)
        alpha = max(0.0, min(1.0, alpha))
        target_fraction = target_start + alpha * (target_end - target_start)
        lower_fraction = max(0.0, target_fraction - hydro_band_fraction)
        upper_fraction = min(1.0, target_fraction + hydro_band_fraction)

        if lower_fraction > upper_fraction:
            raise ValueError(
                f"Banda hydro inválida en {window_label}: "
                f"lower={lower_fraction:.6f}, upper={upper_fraction:.6f}"
            )

        trajectory[window_label] = {
            "constraint_snapshot": constraint_snapshot,
            "window_end": window_end,
            "target_fraction": float(target_fraction),
            "lower_fraction": float(lower_fraction),
            "upper_fraction": float(upper_fraction),
        }

    return trajectory


def add_hydro_terminal_band_constraint(
    grid,
    snapshots,
    lower_fraction: float,
    upper_fraction: float,
    constraint_name="rolling_hydro_terminal_soc_band",
):
    # Rolling horizon: banda terminal agregada para stores hydro.
    add_hydro_terminal_band_constraint_at_snapshot(
        grid,
        snapshot=snapshots[-1],
        lower_fraction=lower_fraction,
        upper_fraction=upper_fraction,
        constraint_name=constraint_name,
    )


def add_hydro_terminal_band_constraint_at_snapshot(
    grid,
    snapshot,
    lower_fraction: float,
    upper_fraction: float,
    constraint_name="hydro_terminal_soc_band",
):
    # Banda terminal agregada para stores hydro en un snapshot arbitrario.
    lower_fraction = float(lower_fraction)
    upper_fraction = float(upper_fraction)

    if lower_fraction < 0.0 or upper_fraction > 1.0:
        raise ValueError(
            "Las fracciones terminales hydro deben estar dentro del rango [0, 1]."
        )

    if lower_fraction > upper_fraction:
        raise ValueError(
            f"Banda terminal hydro inválida: lower={lower_fraction}, upper={upper_fraction}"
        )

    m = grid.model

    if "Store-e" not in m.variables:
        warnings.warn("No existe la variable Store-e. No se añade banda terminal hydro.")
        return

    hydro_stores = grid.stores.index[grid.stores.carrier == "hydro"]

    if len(hydro_stores) == 0:
        print("No hay stores hydro. No se añade banda terminal hydro.")
        return

    target_snapshot = pd.Timestamp(snapshot)
    hydro_soc_final = m.variables["Store-e"].loc[target_snapshot, hydro_stores]
    hydro_e_nom = pd.to_numeric(
        grid.stores.loc[hydro_stores, "e_nom"],
        errors="coerce",
    ).fillna(0.0)
    hydro_e_nom_total = float(hydro_e_nom.sum())

    if hydro_e_nom_total <= 0.0:
        warnings.warn(
            "La capacidad energética total de stores hydro es cero o no válida. "
            "No se añade banda terminal hydro."
        )
        return

    m.add_constraints(
        hydro_soc_final.sum() >= lower_fraction * hydro_e_nom_total,
        name=f"{constraint_name}_min",
    )
    m.add_constraints(
        hydro_soc_final.sum() <= upper_fraction * hydro_e_nom_total,
        name=f"{constraint_name}_max",
    )

    print(
        "Restricción terminal hydro añadida "
        f"para {len(hydro_stores)} stores; "
        f"snapshot={target_snapshot}, "
        f"banda [{lower_fraction:.4f}, {upper_fraction:.4f}].",
        flush=True,
    )


def add_phs_terminal_soc_constraint(
    grid,
    snapshots,
    battery_specs: list[dict] | None = None,
    constraint_name="rolling_phs_terminal_soc",
):
    # Rolling horizon: para PHS con Cyclic SOC input = 1, exigir SOC_final >= SOC_inicial.
    if grid.stores.empty:
        print("Rolling horizon: no hay stores. No se añade restricción terminal PHS.")
        return

    m = grid.model

    if "Store-e" not in m.variables:
        warnings.warn("No existe la variable Store-e. No se añade restricción terminal PHS.")
        return

    phs_stores = grid.stores.index[grid.stores.carrier == "PHS"]

    if len(phs_stores) == 0:
        print("Rolling horizon: no hay PHS. Restricciones terminales PHS añadidas: 0.")
        return

    if battery_specs is None:
        warnings.warn(
            "No se recibieron battery_specs para filtrar PHS por Cyclic SOC (0/1). "
            "No se añade restricción terminal PHS."
        )
        return

    cyclic_phs_stores = {
        spec["store_name"]
        for spec in battery_specs
        if spec.get("carrier") == "PHS" and bool(spec.get("cyclic_soc_input", False))
    }
    constrained_stores = [store for store in phs_stores if store in cyclic_phs_stores]

    if not constrained_stores:
        print("Rolling horizon: PHS con Cyclic SOC (0/1)=1: 0.")
        return

    if "e_initial" not in grid.stores.columns:
        warnings.warn(
            "grid.stores no contiene e_initial. No se añade restricción terminal PHS."
        )
        return

    final_snapshot = snapshots[-1]
    store_e = m.variables["Store-e"]
    added = 0

    for store_name in constrained_stores:
        initial_soc = pd.to_numeric(
            pd.Series([grid.stores.loc[store_name, "e_initial"]]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(initial_soc):
            warnings.warn(
                f"SOC inicial no válido para PHS '{store_name}'. "
                "No se añade su restricción terminal."
            )
            continue

        m.add_constraints(
            store_e.loc[final_snapshot, store_name] >= float(initial_soc),
            name=f"{constraint_name}_{_constraint_name_suffix(store_name)}",
        )
        added += 1

    print(
        "Rolling horizon: restricciones PHS SOC_final >= SOC_inicial añadidas: "
        f"{added}.",
        flush=True,
    )


def add_intermediate_hydro_trajectory_constraints(
    grid,
    hydro_soc_target_trajectory: dict,
):
    if not hydro_soc_target_trajectory:
        return

    for window_label, target in hydro_soc_target_trajectory.items():
        print(
            "Intermediate hydro constraint | "
            f"{window_label}: "
            f"constraint_snapshot={target['constraint_snapshot']}, "
            f"window_end={target['window_end']}, "
            f"target={target['target_fraction']:.4f}, "
            f"lower={target['lower_fraction']:.4f}, "
            f"upper={target['upper_fraction']:.4f}",
            flush=True,
        )
        add_hydro_terminal_band_constraint_at_snapshot(
            grid,
            snapshot=target["constraint_snapshot"],
            lower_fraction=target["lower_fraction"],
            upper_fraction=target["upper_fraction"],
            constraint_name=f"{window_label}_intermediate_hydro_terminal_soc_band",
        )


def add_intermediate_phs_terminal_constraints(
    grid,
    block_windows,
    battery_specs: list[dict] | None = None,
    constraint_name="intermediate_phs_terminal_soc",
):
    # Simulación anual: para PHS cíclicos, SOC al final del bloque >= SOC al inicio.
    if grid.stores.empty:
        print("Intermediate storage constraints: no hay stores para PHS.")
        return

    m = grid.model

    if "Store-e" not in m.variables:
        warnings.warn(
            "No existe la variable Store-e. No se añaden restricciones intermedias PHS."
        )
        return

    phs_stores = grid.stores.index[grid.stores.carrier == "PHS"]

    if len(phs_stores) == 0:
        print("Intermediate storage constraints: no hay PHS. Restricciones PHS añadidas: 0.")
        return

    if battery_specs is None:
        warnings.warn(
            "No se recibieron battery_specs para filtrar PHS por Cyclic SOC (0/1). "
            "No se añaden restricciones intermedias PHS."
        )
        return

    cyclic_phs_stores = {
        spec["store_name"]
        for spec in battery_specs
        if spec.get("carrier") == "PHS" and bool(spec.get("cyclic_soc_input", False))
    }
    constrained_stores = [store for store in phs_stores if store in cyclic_phs_stores]

    if not constrained_stores:
        print("Intermediate storage constraints: PHS con Cyclic SOC (0/1)=1: 0.")
        return

    store_e = m.variables["Store-e"]
    added = 0

    for block_number, block_snapshots in enumerate(block_windows, start=1):
        block_snapshots = pd.DatetimeIndex(block_snapshots)
        if block_snapshots.empty:
            continue

        start_snapshot = pd.Timestamp(block_snapshots[0])
        end_snapshot = pd.Timestamp(block_snapshots[-1])

        for store_name in constrained_stores:
            m.add_constraints(
                store_e.loc[end_snapshot, store_name] >= store_e.loc[start_snapshot, store_name],
                name=(
                    f"{constraint_name}_{block_number:03d}_"
                    f"{_constraint_name_suffix(store_name)}"
                ),
            )
            added += 1

    print(
        "Intermediate storage constraints: restricciones PHS añadidas; "
        f"n_PHS={len(constrained_stores)}, "
        f"n_blocks={len(block_windows)}, "
        f"n_constraints={added}.",
        flush=True,
    )


def add_intermediate_storage_terminal_constraints(
    grid,
    block_windows,
    hydro_soc_target_trajectory: dict,
    battery_specs: list[dict] | None,
):
    add_intermediate_hydro_trajectory_constraints(
        grid,
        hydro_soc_target_trajectory,
    )
    add_intermediate_phs_terminal_constraints(
        grid,
        block_windows,
        battery_specs=battery_specs,
    )


def add_batterystore_residual_value_to_objective(
    grid,
    snapshots,
    residual_value_eur_per_mwh: float,
):
    # Rolling horizon: valor residual de SOC final BatteryStore en el objetivo.
    residual_value_eur_per_mwh = float(residual_value_eur_per_mwh or 0.0)

    if residual_value_eur_per_mwh <= 0.0:
        return

    m = grid.model

    if "Store-e" not in m.variables:
        warnings.warn(
            "No existe la variable Store-e. No se añade valor residual BatteryStore."
        )
        return

    batterystore_stores = grid.stores.index[grid.stores.carrier == "BatteryStore"]

    if len(batterystore_stores) == 0:
        warnings.warn(
            "Valor residual BatteryStore > 0, pero no hay stores con carrier "
            "'BatteryStore'. No se modifica el objetivo."
        )
        return

    final_snapshot = snapshots[-1]
    batterystore_soc_final = m.variables["Store-e"].loc[
        final_snapshot,
        batterystore_stores,
    ]
    residual_term = residual_value_eur_per_mwh * batterystore_soc_final.sum()

    m.objective += -residual_term

    print(
        "Rolling horizon: valor residual BatteryStore añadido al objetivo; "
        f"n_BatteryStore={len(batterystore_stores)}, "
        f"residual_value={residual_value_eur_per_mwh:.2f} €/MWh.",
        flush=True,
    )


def get_batterystore_min_final_soc_requirement(
    grid,
    min_final_soc_fraction: float,
) -> tuple[int, float]:
    min_final_soc_fraction = float(min_final_soc_fraction or 0.0)

    if min_final_soc_fraction <= 0.0:
        return 0, 0.0

    if not hasattr(grid, "stores") or grid.stores.empty:
        return 0, 0.0

    batterystore_stores = grid.stores.index[grid.stores.carrier == "BatteryStore"]

    if len(batterystore_stores) == 0:
        return 0, 0.0

    e_nom = pd.to_numeric(
        grid.stores.loc[batterystore_stores, "e_nom"],
        errors="coerce",
    ).fillna(0.0)
    min_final_soc_mwh_total = float((min_final_soc_fraction * e_nom).sum())

    return len(batterystore_stores), min_final_soc_mwh_total


def add_batterystore_min_final_soc_constraint(
    grid,
    snapshots,
    min_final_soc_fraction: float,
    constraint_name="rolling_batterystore_min_final_soc",
):
    # Rolling horizon: SOC final mínimo individual para BatteryStore.
    min_final_soc_fraction = float(min_final_soc_fraction or 0.0)

    if min_final_soc_fraction <= 0.0:
        return

    if min_final_soc_fraction > 0.95:
        raise ValueError(
            "rolling_batterystore_min_final_soc_fraction debe ser <= 0.95."
        )

    m = grid.model

    if "Store-e" not in m.variables:
        warnings.warn(
            "No existe la variable Store-e. No se añade restricción "
            "BatteryStore minimum final SOC."
        )
        return

    batterystore_stores = grid.stores.index[grid.stores.carrier == "BatteryStore"]

    if len(batterystore_stores) == 0:
        warnings.warn(
            "BatteryStore minimum final SOC > 0, pero no hay stores con "
            "carrier 'BatteryStore'. No se añade restricción."
        )
        return

    final_snapshot = snapshots[-1]
    store_e = m.variables["Store-e"]
    added = 0
    min_final_soc_mwh_total = 0.0

    for store_name in batterystore_stores:
        e_nom = pd.to_numeric(
            pd.Series([grid.stores.loc[store_name, "e_nom"]]),
            errors="coerce",
        ).iloc[0]

        if pd.isna(e_nom) or float(e_nom) <= 0.0:
            warnings.warn(
                f"Capacidad e_nom no válida para BatteryStore '{store_name}'. "
                "No se añade su restricción de SOC final mínimo."
            )
            continue

        min_final_soc_mwh = min_final_soc_fraction * float(e_nom)
        m.add_constraints(
            store_e.loc[final_snapshot, store_name] >= min_final_soc_mwh,
            name=f"{constraint_name}_{_constraint_name_suffix(store_name)}",
        )
        added += 1
        min_final_soc_mwh_total += min_final_soc_mwh

    print(
        "Rolling horizon: BatteryStore minimum final SOC constraint added; "
        f"min_final_soc={min_final_soc_fraction:.2%}, "
        f"n_BatteryStore={added}, "
        f"min_final_SOC_total={min_final_soc_mwh_total:.4f} MWh.",
        flush=True,
    )


def get_final_batterystore_soc_mwh(grid) -> tuple[int, float]:
    if not hasattr(grid, "stores") or grid.stores.empty:
        return 0, 0.0

    batterystore_stores = grid.stores.index[grid.stores.carrier == "BatteryStore"]

    if len(batterystore_stores) == 0:
        return 0, 0.0

    if (
        not hasattr(grid, "stores_t")
        or not hasattr(grid.stores_t, "e")
        or grid.stores_t.e.empty
    ):
        warnings.warn(
            "No se encontró grid.stores_t.e para calcular SOC final BatteryStore."
        )
        return len(batterystore_stores), 0.0

    final_snapshot = grid.snapshots[-1]
    available_stores = [
        store for store in batterystore_stores
        if store in grid.stores_t.e.columns
    ]

    missing_stores = set(batterystore_stores) - set(available_stores)
    for store_name in sorted(missing_stores):
        warnings.warn(
            f"No se encontró SOC final para BatteryStore '{store_name}'."
        )

    if not available_stores:
        return len(batterystore_stores), 0.0

    final_soc = pd.to_numeric(
        grid.stores_t.e.loc[final_snapshot, available_stores],
        errors="coerce",
    ).fillna(0.0)

    return len(available_stores), float(final_soc.sum())

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_default_input_file() -> Path:
    return get_base_dir() / "GridInputs.xlsx"


def parse_bool_setting(value) -> bool:
    if isinstance(value, str):
        return value.strip().lower() in ["true", "1", "yes", "sí", "si"]
    return bool(value)


def validate_rolling_horizon_days(value) -> int:
    if isinstance(value, float) and not value.is_integer():
        raise ValueError("rolling_horizon_days debe ser un entero entre 1 y 7.")

    try:
        rolling_horizon_days = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError("rolling_horizon_days debe ser un entero entre 1 y 7.") from exc

    if rolling_horizon_days < 1 or rolling_horizon_days > 7:
        raise ValueError("rolling_horizon_days debe estar entre 1 y 7 días.")

    return rolling_horizon_days


def validate_rolling_hydro_soc_band_percent(value) -> float:
    if value is None:
        value = 5.0

    try:
        band_percent = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "rolling_hydro_soc_band_percent debe ser un número entre 0 y 30."
        ) from exc

    if band_percent < 0.0 or band_percent > 30.0:
        raise ValueError(
            "rolling_hydro_soc_band_percent debe estar entre 0 y 30%."
        )

    return band_percent


def validate_intermediate_storage_constraint_days(value) -> int:
    if isinstance(value, float) and not value.is_integer():
        raise ValueError(
            "intermediate_storage_constraint_days debe ser un entero entre 1 y 7."
        )

    try:
        intermediate_days = int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "intermediate_storage_constraint_days debe ser un entero entre 1 y 7."
        ) from exc

    if intermediate_days < 1 or intermediate_days > 7:
        raise ValueError(
            "intermediate_storage_constraint_days debe estar entre 1 y 7 días."
        )

    return intermediate_days


def validate_intermediate_hydro_soc_band_percent(value) -> float:
    if value is None:
        value = 0.5

    try:
        band_percent = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "intermediate_hydro_soc_band_percent debe ser un número entre 0 y 30."
        ) from exc

    if band_percent < 0.0 or band_percent > 30.0:
        raise ValueError(
            "intermediate_hydro_soc_band_percent debe estar entre 0 y 30%."
        )

    return band_percent


def validate_rolling_batterystore_residual_value(value) -> float:
    if value is None:
        value = 0.0

    if isinstance(value, str) and value.strip() == "":
        value = 0.0

    try:
        residual_value = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "rolling_batterystore_residual_value_eur_per_mwh debe ser un número no negativo."
        ) from exc

    if residual_value < 0.0:
        raise ValueError(
            "rolling_batterystore_residual_value_eur_per_mwh no puede ser negativo."
        )

    return residual_value


def validate_rolling_batterystore_min_final_soc_percent(value) -> float:
    if value is None:
        value = 0.0

    if isinstance(value, str) and value.strip() == "":
        value = 0.0

    try:
        min_final_soc_percent = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(
            "rolling_batterystore_min_final_soc_percent debe ser un número entre 0 y 95."
        ) from exc

    if min_final_soc_percent < 0.0 or min_final_soc_percent > 95.0:
        raise ValueError(
            "rolling_batterystore_min_final_soc_percent debe estar entre 0 y 95%."
        )

    return min_final_soc_percent


def storage_optimization_enabled(df_StorageUnit: pd.DataFrame) -> bool:
    if df_StorageUnit is None or df_StorageUnit.empty:
        return False

    if "Optimization mode" not in df_StorageUnit.columns:
        return False

    modes = (
        df_StorageUnit["Optimization mode"]
        .fillna("Fixed")
        .astype(str)
        .str.strip()
    )

    return bool(modes.isin(["Optimize both", "Optimize MW", "Optimize MWh"]).any())


def generate_rolling_windows(snapshots, rolling_horizon_days) -> list[pd.DatetimeIndex]:
    rolling_horizon_days = validate_rolling_horizon_days(rolling_horizon_days)
    snapshots = pd.DatetimeIndex(snapshots)

    if snapshots.empty:
        raise ValueError("No hay snapshots para generar ventanas rolling horizon.")

    if snapshots.has_duplicates:
        raise ValueError("El horizonte largo contiene snapshots duplicados.")

    window_size = rolling_horizon_days * 24
    windows = [
        pd.DatetimeIndex(snapshots[start:start + window_size])
        for start in range(0, len(snapshots), window_size)
    ]

    windows = [window for window in windows if not window.empty]

    covered_snapshots = pd.DatetimeIndex([])
    if windows:
        covered_snapshots = windows[0]
        for window in windows[1:]:
            covered_snapshots = covered_snapshots.append(window)

    if not covered_snapshots.equals(snapshots):
        raise ValueError("Las ventanas rolling horizon no cubren exactamente el horizonte largo.")

    return windows


def _as_optional_float(value):
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() in ["", "none", "nan"]:
        return None
    return float(value)


def _as_optional_int(value, zero_is_none=False):
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() in ["", "none", "nan"]:
        return None

    value = int(float(value))

    if zero_is_none and value == 0:
        return None

    return value


def _get_solver_run_options(params: pd.Series) -> dict:
    line_flow_penalty = _as_optional_float(params.get("line_flow_penalty"))
    if line_flow_penalty is None:
        line_flow_penalty = 0.0

    use_line_length_scaling_raw = params.get("use_line_length_scaling", True)
    use_line_length_scaling = parse_bool_setting(use_line_length_scaling_raw)

    return {
        "solver_name": str(params.get("solver_name")),
        "mip_rel_gap": _as_optional_float(params.get("mip_rel_gap")),
        "time_limit": _as_optional_int(params.get("time_limit"), zero_is_none=True),
        "threads": _as_optional_int(params.get("threads"), zero_is_none=False),
        "mip_focus": _as_optional_int(params.get("mip_focus")),
        "method": _as_optional_int(params.get("method")),
        "bar_homogeneous": _as_optional_int(params.get("bar_homogeneous"), zero_is_none=False),
        "crossover": _as_optional_int(params.get("crossover")),
        "numeric_focus": _as_optional_int(params.get("numeric_focus")),
        "bar_conv_tol": _as_optional_float(params.get("bar_conv_tol")),
        "feasibility_tol": _as_optional_float(params.get("feasibility_tol")),
        "optimality_tol": _as_optional_float(params.get("optimality_tol")),
        "line_flow_penalty": line_flow_penalty,
        "use_line_length_scaling": use_line_length_scaling,
    }


def _slice_indexed_dataframe(
    df: pd.DataFrame,
    snapshots: pd.DatetimeIndex,
    name: str,
) -> pd.DataFrame:
    result = df.copy()
    result.index = pd.to_datetime(result.index)
    result = result.sort_index()

    missing = pd.DatetimeIndex(snapshots).difference(result.index)
    if len(missing) > 0:
        raise ValueError(
            f"Faltan {len(missing)} snapshots en {name}; primer missing: {missing[0]}"
        )

    return result.loc[snapshots].copy()


def _slice_time_column_dataframe(
    df: pd.DataFrame,
    snapshots: pd.DatetimeIndex,
    name: str,
) -> pd.DataFrame:
    result = df.copy()

    if "time" in result.columns:
        result["time"] = pd.to_datetime(result["time"]).dt.round("h")
        result = result.set_index("time").sort_index()
    elif "snapshot" in result.columns:
        result["snapshot"] = pd.to_datetime(result["snapshot"]).dt.round("h")
        result = result.set_index("snapshot").sort_index()
    else:
        # Rolling horizon: algunos pasos pueden recibir perfiles con "time"
        # ya convertido en índice. Aceptamos esa forma sin modificar valores.
        if not isinstance(result.index, pd.DatetimeIndex):
            raise KeyError(
                f"{name} debe tener una columna 'time' o un índice temporal."
            )

        result.index = pd.DatetimeIndex(pd.to_datetime(result.index)).round("h")

        result = result.sort_index()

    result.index.name = "time"

    missing = pd.DatetimeIndex(snapshots).difference(result.index)
    if len(missing) > 0:
        raise ValueError(
            f"Faltan {len(missing)} snapshots en {name}; primer missing: {missing[0]}"
        )

    window = result.loc[snapshots].copy()
    window.index.name = "time"
    return window.reset_index()


def _build_available_renewable(
    df_solar_node_profiles: pd.DataFrame,
    df_wind_node_profiles: pd.DataFrame,
) -> pd.DataFrame:
    def with_time_index(df: pd.DataFrame, name: str) -> pd.DataFrame:
        result = df.copy()

        if "time" in result.columns:
            result["time"] = pd.to_datetime(result["time"]).dt.round("h")
            result = result.set_index("time")
        elif "snapshot" in result.columns:
            result["snapshot"] = pd.to_datetime(result["snapshot"]).dt.round("h")
            result = result.set_index("snapshot")
        else:
            # Rolling horizon: después de cortar ventanas, "time" puede
            # llegar ya como índice. No hacemos set_index dos veces.
            if not isinstance(result.index, pd.DatetimeIndex):
                raise KeyError(
                    f"{name} debe tener una columna 'time', una columna 'snapshot' "
                    "o un índice temporal."
                )

            result.index = pd.DatetimeIndex(pd.to_datetime(result.index)).round("h")

        result.index.name = "snapshot"
        return result.copy()

    df_solar_available = with_time_index(df_solar_node_profiles, "df_solar_node_profiles")
    df_wind_available = with_time_index(df_wind_node_profiles, "df_wind_node_profiles")

    df_solar_available.columns = [f"PV_{c}" for c in df_solar_available.columns]
    df_wind_available.columns = [f"Wind_{c}" for c in df_wind_available.columns]

    df_available_renewable = pd.concat(
        [df_solar_available, df_wind_available],
        axis=1,
    )

    df_available_renewable.index = pd.to_datetime(df_available_renewable.index)
    df_available_renewable.index.name = "snapshot"

    return df_available_renewable


def _grid_inputs_has_morocco_connection(df_Grid_connection: pd.DataFrame) -> bool:
    if df_Grid_connection is None or df_Grid_connection.empty:
        return False

    df_Grid_connection = df_Grid_connection.copy()
    df_Grid_connection.columns = df_Grid_connection.columns.astype(str).str.strip()

    for col in ["PCC name", "PCC"]:
        if col in df_Grid_connection.columns:
            values = df_Grid_connection[col].dropna().astype(str)
            if values.str.contains("Morocco", case=False, regex=False).any():
                return True

    return False


def _read_morocco_exchange_series(
    base_dir: Path,
    full_snapshots: pd.DatetimeIndex,
) -> pd.Series:
    lista_df = []

    for n in range(1, 12):
        nombre_doc = f"saldo_marruecos{n}.xls"
        ruta_doc = base_dir / "System_data" / "Saldo_Marruecos_ESIOS" / nombre_doc
        tablas = pd.read_html(ruta_doc, header=0, decimal=",", thousands=".")
        df = tablas[0]
        df = df[["name", "value", "datetime"]]
        df = df[df["name"] == "Generación programada P48 Saldo Marruecos"]
        df = df[["datetime", "value"]]
        df["datetime"] = pd.to_datetime(df["datetime"], utc=True)
        df["datetime"] = df["datetime"].dt.tz_convert("Europe/Madrid").dt.tz_localize(None)
        df = df.set_index("datetime")
        df["value"] = pd.to_numeric(df["value"], errors="coerce")
        lista_df.append(df)

    df_final = pd.concat(lista_df)
    df_final = df_final.sort_index()
    df_final = df_final[~df_final.index.duplicated(keep="first")]

    saldo = df_final["value"].reindex(full_snapshots)

    missing = saldo.isna().sum()
    total = len(saldo)
    print(
        "Datos faltantes Marruecos rolling horizon después de reindexar: "
        f"{missing}/{total} ({missing/total:.1%})"
    )

    saldo = saldo.interpolate(method="time", limit=3)

    missing = saldo.isna().sum()
    total = len(saldo)
    print(
        "Datos faltantes Marruecos rolling horizon después de interpolar: "
        f"{missing}/{total} ({missing/total:.1%})"
    )

    return saldo.fillna(0)


def _has_storage_rows(df_StorageUnit: pd.DataFrame) -> bool:
    return (
        df_StorageUnit is not None
        and not df_StorageUnit.empty
        and "LOCATION" in df_StorageUnit.columns
        and pd.notna(df_StorageUnit.loc[df_StorageUnit.index[0], "LOCATION"])
    )


def _window_duration_days(window_snapshots: pd.DatetimeIndex) -> int:
    if len(window_snapshots) == 0:
        raise ValueError("La ventana rolling horizon no contiene snapshots.")

    if len(window_snapshots) % 24 != 0:
        raise ValueError(
            "Las ventanas rolling horizon deben contener días completos "
            f"en este modelo horario. Snapshots recibidos: {len(window_snapshots)}."
        )

    return len(window_snapshots) // 24


def _build_window_system_parameters(
    system_parameters: dict,
    window_snapshots: pd.DatetimeIndex,
    window_label: str,
) -> dict:
    window_days = _window_duration_days(window_snapshots)
    window_parameters = dict(system_parameters)
    window_parameters["Static / Multiperiod"] = "Multiperiod"
    window_parameters["Start date (dd/mm/aaaa)"] = window_snapshots[0].to_pydatetime()
    window_parameters["Simulation duration (days)"] = window_days

    notes = str(window_parameters.get("Notes", "")).strip()
    window_parameters["Notes"] = (
        f"{notes} | {window_label}" if notes else window_label
    )

    return window_parameters


def apply_initial_storage_state(model_inputs: dict, storage_state: dict | None) -> dict:
    # Rolling horizon: inyecta el SOC final de la ventana anterior.
    updated = dict(model_inputs)
    updated["initial_storage_state"] = storage_state or None
    return updated


def apply_initial_generator_power(
    model_inputs: dict,
    generator_power_state: dict | None,
) -> dict:
    # Rolling horizon: inyecta la potencia final de generadores con rampas.
    updated = dict(model_inputs)
    updated["initial_generator_power"] = generator_power_state or None
    return updated


def extract_final_storage_state(results) -> dict:
    grid = results["grid"] if isinstance(results, dict) else results
    storage_state = {}

    if not hasattr(grid, "stores") or grid.stores.empty:
        return storage_state

    if (
        not hasattr(grid, "stores_t")
        or not hasattr(grid.stores_t, "e")
        or grid.stores_t.e.empty
    ):
        warnings.warn(
            "No se encontró grid.stores_t.e para extraer SOC final del rolling horizon."
        )
        return storage_state

    final_snapshot = grid.snapshots[-1]

    for store_name in grid.stores.index:
        if store_name not in grid.stores_t.e.columns:
            warnings.warn(
                f"No se encontró resultado de SOC para el almacenamiento '{store_name}'."
            )
            continue

        storage_state[store_name] = float(grid.stores_t.e.loc[final_snapshot, store_name])

    return storage_state


def extract_final_generator_power(results) -> dict:
    grid = results["grid"] if isinstance(results, dict) else results
    generator_power_state = {}

    if not hasattr(grid, "generators") or grid.generators.empty:
        return generator_power_state

    if (
        not hasattr(grid, "generators_t")
        or not hasattr(grid.generators_t, "p")
        or grid.generators_t.p.empty
    ):
        warnings.warn(
            "No se encontró grid.generators_t.p para extraer potencias finales."
        )
        return generator_power_state

    ramped = pd.Series(False, index=grid.generators.index)

    for col in ["ramp_limit_up", "ramp_limit_down"]:
        if col in grid.generators.columns:
            ramped = ramped | grid.generators[col].notna()

    for col in ["ramp_limit_start_up", "ramp_limit_shut_down"]:
        if col in grid.generators.columns:
            values = pd.to_numeric(grid.generators[col], errors="coerce")
            ramped = ramped | (values.notna() & ((values - 1.0).abs() > 1e-12))

    ramped_generators = grid.generators.index[ramped]
    final_snapshot = grid.snapshots[-1]

    for generator_name in ramped_generators:
        if generator_name not in grid.generators_t.p.columns:
            warnings.warn(
                f"No se encontró resultado de potencia para el generador '{generator_name}'."
            )
            continue

        generator_power_state[generator_name] = float(
            grid.generators_t.p.loc[final_snapshot, generator_name]
        )

    return generator_power_state


def _normalize_temporal_result_table(
    table,
    key: str,
    window_result: dict,
) -> pd.DataFrame:
    if isinstance(table, pd.Series):
        df = table.to_frame()
    else:
        df = table.copy()

    if "snapshot" in df.columns:
        df = df.set_index("snapshot")
    elif "time" in df.columns:
        df = df.set_index("time")
    elif not isinstance(df.index, pd.DatetimeIndex):
        window_snapshots = pd.DatetimeIndex(window_result["grid"].snapshots)

        if len(df) != len(window_snapshots):
            raise ValueError(
                f"No se puede normalizar el índice temporal de {key} en "
                f"{window_result['window_label']}: len(df)={len(df)}, "
                f"len(window_snapshots)={len(window_snapshots)}."
            )

        warnings.warn(
            f"{key} en {window_result['window_label']} no tenía índice temporal; "
            "se usan los snapshots reales de la ventana."
        )
        df.index = window_snapshots

    df.index = pd.DatetimeIndex(pd.to_datetime(df.index))

    if df.index.isna().any():
        raise ValueError(
            f"El índice temporal de {key} en {window_result['window_label']} contiene NaT."
        )

    df.index.name = "snapshot"
    return df


def _prepare_rolling_horizon_inputs(
    input_path: Path,
    data: dict,
    system_parameters: dict,
    df_SYS_settings: pd.DataFrame,
) -> dict:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon != "Multiperiod":
        raise ValueError("Rolling horizon solo está disponible para simulaciones Multiperiod.")

    if storage_optimization_enabled(data["StorageUnit"]):
        raise ValueError(
            "Rolling horizon no está disponible cuando StorageUnit contiene "
            "modos de optimización de batería/almacenamiento. Usa capacidades fijas."
        )

    startdate = params["Start date (dd/mm/aaaa)"]
    duration = int(params["Simulation duration (days)"])
    full_snapshots = pd.date_range(startdate, periods=duration * 24, freq="h")

    BASE_DIR = Path(__file__).resolve().parent

    # -- TRATAMIENTO DE PERFILES DE DEMANDA de ESPAÑA --
    rutaDemandaXccaa = BASE_DIR / "System_data" / "demanda_anual_ccaa_REE_2015_2025.xlsx"
    df_consumos_anuales_CCAA = pd.read_excel(
        rutaDemandaXccaa,
        sheet_name="demanda_ccaa_largo",
        usecols="A,B,E",
    )
    df_consumos_anuales_CCAA = df_consumos_anuales_CCAA.dropna()

    demr_folder = BASE_DIR / "System_data" / "DemandaDiariaSistemaElectricoPeninsular"
    df_demanda_ccaa = build_hourly_demand_by_region(
        df_consumos_anuales_CCAA=df_consumos_anuales_CCAA,
        demr_folder=demr_folder,
        startdate=startdate,
        days=duration,
    )

    ruta2 = BASE_DIR / "System_data" / "2013PyPSA_Network.xlsx"
    df_PyPSA2013_load_profiles = pd.read_excel(
        ruta2,
        sheet_name="loads_timeseries",
    )

    df_monthly_node_weights_ESP = build_monthly_nodal_load_weights_ES(
        df_pypsa_load_profiles=df_PyPSA2013_load_profiles,
        node_to_region=node_to_region,
        exclude_portugal=True,
    )

    df_demanda_nodal_ESP = build_hourly_nodal_demand(
        df_demanda_ccaa=df_demanda_ccaa,
        df_monthly_node_weights=df_monthly_node_weights_ESP,
    )
    df_demanda_nodal_ESP = df_demanda_nodal_ESP.set_index("time")

    # -- TRATAMIENTO DE PERFILES DE DEMANDA de PORTUGAL --
    ruta_demanda_PT = BASE_DIR / "System_data" / "DemandaPT.xlsx"
    df_total_demand_pt = pd.read_excel(
        ruta_demanda_PT,
        sheet_name="Portugal demand TS",
        usecols="A:B",
    )
    df_region_demand = pd.read_excel(
        ruta_demanda_PT,
        sheet_name="Consumos mensuales por región",
        usecols="A:F",
        header=1,
    )

    df_demand_regions_hourly = regional_hourly_demand_builder(
        df_regional_demand=df_region_demand,
        df_total_hourly_demand=df_total_demand_pt,
        startdate=startdate,
        days=duration,
    )

    df_monthly_node_weights_PT = build_monthly_nodal_load_weights_PT(
        df_pypsa_load_profiles=df_PyPSA2013_load_profiles,
        node_to_region=node_to_region_PT,
        exclude_spain=True,
    )

    df_demanda_nodal_PT = build_hourly_nodal_demand_PT(
        df_demand_PT_regions=df_demand_regions_hourly,
        df_monthly_node_weights=df_monthly_node_weights_PT,
    )
    df_demanda_nodal_PT = df_demanda_nodal_PT.set_index("time")

    df_demanda_nodal = pd.concat([df_demanda_nodal_PT, df_demanda_nodal_ESP], axis=1)
    df_demanda_nodal.index = pd.to_datetime(df_demanda_nodal.index)

    # -- TRATAMIENTO DE PERFILES DE RENOVABLES --
    df_solar_instaled_capacity = pd.read_excel(
        input_path,
        engine="openpyxl",
        sheet_name="Gen_PV_and_Wind",
        usecols="G:AB",
        skiprows=2,
        nrows=14,
    )
    df_wind_instaled_capacity = pd.read_excel(
        input_path,
        engine="openpyxl",
        sheet_name="Gen_PV_and_Wind",
        usecols="G:AB",
        skiprows=19,
        nrows=14,
    )

    df_wind_instaled_capacity.columns = (
        df_wind_instaled_capacity.columns
        .str.replace(r"\.\d+$", "", regex=True)
        .str.strip()
    )
    df_solar_instaled_capacity.columns = (
        df_solar_instaled_capacity.columns
        .str.replace(r"\.\d+$", "", regex=True)
        .str.strip()
    )

    rutaTSRenewablesninja = BASE_DIR / "System_data" / "RenewablesNinja_Time_Series.xlsx"
    df_solar_profiles = pd.read_excel(
        rutaTSRenewablesninja,
        sheet_name="TS_PV_Profiles",
        usecols="A:W",
    )
    df_wind_profiles = pd.read_excel(
        rutaTSRenewablesninja,
        sheet_name="TS_Wind_Profiles",
        usecols="A:W",
    )
    CFwind = df_wind_profiles.drop("time", axis=1).mean().mean()
    CFsolar = df_solar_profiles.drop("time", axis=1).mean().mean()

    df_node_weights = pd.read_excel(
        input_path,
        sheet_name="Gen_PV_and_Wind",
        usecols="B:E",
        skiprows=2,
        nrows=97,
    )
    df_solar_node_weights = df_node_weights[df_node_weights["Renewable Type"] == "PV"]
    df_wind_node_weights = df_node_weights[df_node_weights["Renewable Type"] == "Wind"]

    df_solar_node_profiles = renewable_profile_builder(
        df_installed_capacity=df_solar_instaled_capacity,
        df_node_weights=df_solar_node_weights,
        df_profiles=df_solar_profiles,
        days=duration,
        startdate=startdate,
    )

    df_wind_node_profiles = renewable_profile_builder(
        df_installed_capacity=df_wind_instaled_capacity,
        df_node_weights=df_wind_node_weights,
        df_profiles=df_wind_profiles,
        days=duration,
        startdate=startdate,
    )

    # -- LECTURA DE PRECIOS DE FRANCIA --
    rutaFrancia = BASE_DIR / "System_data" / "precios_francia_2015_2024.xlsx"
    df_TS_Energy_Prices = pd.read_excel(
        rutaFrancia,
        sheet_name="PreciosFrancia",
        usecols="A:B",
    )
    df_TS_Energy_Prices["time"] = pd.to_datetime(df_TS_Energy_Prices["time"]).dt.round("h")
    df_TS_Energy_Prices = df_TS_Energy_Prices.set_index("time")
    df_TS_Energy_Prices["Precio Francia (€/MWh)"] = (
        df_TS_Energy_Prices["Precio Francia (€/MWh)"].clip(lower=0)
    )

    morocco_exchange_series = None
    if _grid_inputs_has_morocco_connection(data["Grid_connection"]):
        morocco_exchange_series = _read_morocco_exchange_series(
            BASE_DIR,
            full_snapshots,
        )

    # -- HIDRO, ROR Y EMBALSES --
    df_hydro_inflow_scaled = build_hydro_inflow(
        base_dir=BASE_DIR,
        startdate=startdate,
        days=duration,
    )

    df_ror_p_max_pu_scaled = build_ror_p_max_pu(
        base_dir=BASE_DIR,
        startdate=startdate,
        days=duration,
    )

    ruta_llenado_embalses = BASE_DIR / "System_data" / "Embalses.xlsx"
    df_embalses_original = pd.read_excel(ruta_llenado_embalses, usecols="D:G")
    df_embalses_original = df_embalses_original[
        df_embalses_original["ELECTRICO_FLAG"] == 1
    ]

    df_embalses_initial_SOC = get_embalses_closest_date(
        df=df_embalses_original,
        target_date=startdate,
    )

    initial_soc_fraction = (
        df_embalses_initial_SOC["AGUA_ACTUAL"].sum()
        / df_embalses_initial_SOC["AGUA_TOTAL"].sum()
    )
    fecha_final = pd.to_datetime(startdate) + pd.to_timedelta(duration, unit="D")

    economic_settings = build_battery_economic_settings_from_gui(system_parameters)
    if horizon == "Multiperiod" and economic_settings.iloc[0, 0] is not None:
        def crf(i, n):
            if i == 0:
                return 1 / n
            return (i * (1 + i)**n) / ((1 + i)**n - 1)

        interest = economic_settings.iloc[0, 0] / 100
        lifetime = economic_settings.iloc[1, 0]
        CRF = crf(interest, lifetime)
    else:
        CRF = 0

    gas_price, co2_price = CCGT_dataframe_treatment(BASE_DIR, startdate, duration)

    return {
        "BASE_DIR": BASE_DIR,
        "input_path": input_path,
        "data": data,
        "system_parameters": system_parameters,
        "df_SYS_settings_full": df_SYS_settings,
        "full_snapshots": full_snapshots,
        "startdate": startdate,
        "duration": duration,
        "fecha_final": fecha_final,
        "df_demanda_nodal": df_demanda_nodal,
        "df_solar_node_profiles": df_solar_node_profiles,
        "df_wind_node_profiles": df_wind_node_profiles,
        "CFsolar": CFsolar,
        "CFwind": CFwind,
        "df_TS_Energy_Prices": df_TS_Energy_Prices,
        "morocco_exchange_series": morocco_exchange_series,
        "df_hydro_inflow_scaled": df_hydro_inflow_scaled,
        "df_ror_p_max_pu_scaled": df_ror_p_max_pu_scaled,
        "df_embalses_original": df_embalses_original,
        "initial_soc_fraction": initial_soc_fraction,
        "CRF": CRF,
        "gas_price": gas_price,
        "co2_price": co2_price,
        "df_Net_Buses": data["Net_Buses"],
        "df_Net_Lines": data["Net_Lines"],
        "df_Gen_Dispatchable": data["Gen_Dispatchable"],
        "df_StorageUnit": data["StorageUnit"],
        "df_Grid_connection": data["Grid_connection"],
    }


def run_single_rolling_window(
    config: dict,
    window_snapshots,
    previous_storage_state=None,
    previous_generator_power=None,
) -> dict:
    prepared = config["prepared_inputs"]
    window_label = config["window_label"]
    run_timestamp = config["run_timestamp"]
    window_snapshots = pd.DatetimeIndex(window_snapshots)

    window_parameters = _build_window_system_parameters(
        prepared["system_parameters"],
        window_snapshots,
        window_label,
    )

    df_SYS_settings = build_sys_settings_from_gui(window_parameters)
    params = df_SYS_settings["SYSTEM PARAMETERS"]

    grid = build_network(df_SYS_settings)
    if not grid.snapshots.equals(window_snapshots):
        raise ValueError(
            f"Los snapshots de la red no coinciden con {window_label}."
        )

    model_inputs = {}
    model_inputs = apply_initial_storage_state(model_inputs, previous_storage_state)
    model_inputs = apply_initial_generator_power(model_inputs, previous_generator_power)

    df_demanda_nodal = _slice_indexed_dataframe(
        prepared["df_demanda_nodal"],
        window_snapshots,
        "df_demanda_nodal",
    )
    df_solar_node_profiles = _slice_time_column_dataframe(
        prepared["df_solar_node_profiles"],
        window_snapshots,
        "df_solar_node_profiles",
    )
    df_wind_node_profiles = _slice_time_column_dataframe(
        prepared["df_wind_node_profiles"],
        window_snapshots,
        "df_wind_node_profiles",
    )
    df_hydro_inflow_scaled = _slice_indexed_dataframe(
        prepared["df_hydro_inflow_scaled"],
        window_snapshots,
        "df_hydro_inflow_scaled",
    )
    df_ror_p_max_pu_scaled = _slice_indexed_dataframe(
        prepared["df_ror_p_max_pu_scaled"],
        window_snapshots,
        "df_ror_p_max_pu_scaled",
    )

    gas_price = daily_to_snapshots(prepared["gas_price"], grid.snapshots)
    co2_price = daily_to_snapshots(prepared["co2_price"], grid.snapshots)

    add_buses(grid, prepared["df_Net_Buses"])
    add_lines(grid, prepared["df_Net_Lines"])
    add_loads(grid, df_demanda_nodal, df_SYS_settings)

    add_dispatchable_generators(
        grid,
        prepared["df_Gen_Dispatchable"],
        gas_price,
        co2_price,
        df_ror_p_max_pu_scaled,
        initial_generator_power=model_inputs.get("initial_generator_power"),
    )

    add_renewable_generator(
        grid,
        params,
        df_solar_node_profiles,
        df_wind_node_profiles,
    )

    grid_connection(
        grid,
        prepared["df_Grid_connection"],
        prepared["df_TS_Energy_Prices"],
        df_SYS_settings,
        morocco_exchange_series=prepared.get("morocco_exchange_series"),
    )

    if _has_storage_rows(prepared["df_StorageUnit"]):
        battery_specs = add_storage_as_store_links(
            df_SYS_settings,
            grid,
            prepared["df_StorageUnit"],
            prepared["CRF"],
            df_hydro_inflow_scaled,
            prepared["initial_soc_fraction"],
            initial_storage_state=model_inputs.get("initial_storage_state"),
            force_non_cyclic=True,
        )
    else:
        battery_specs = None

    if not grid.storage_units.empty:
        hydro_units = grid.storage_units.index[
            grid.storage_units.carrier == "hydro"
        ]
        hydro_units = [unit for unit in hydro_units if unit in df_hydro_inflow_scaled.columns]
        if hydro_units:
            grid.storage_units_t.inflow.loc[:, hydro_units] = df_hydro_inflow_scaled[hydro_units]

    solver_options = _get_solver_run_options(params)
    hydro_soc_target = config.get("hydro_soc_target")
    batterystore_residual_value = float(
        config.get("rolling_batterystore_residual_value_eur_per_mwh", 0.0)
    )
    batterystore_min_final_soc_percent = float(
        config.get("rolling_batterystore_min_final_soc_percent", 0.0)
    )
    batterystore_min_final_soc_fraction = batterystore_min_final_soc_percent / 100.0
    (
        batterystore_min_final_soc_count,
        batterystore_min_final_soc_mwh_total,
    ) = get_batterystore_min_final_soc_requirement(
        grid,
        batterystore_min_final_soc_fraction,
    )

    if hydro_soc_target is not None:
        print(
            f"{window_label}: terminal hydro en {hydro_soc_target['window_end']} "
            f"target={hydro_soc_target['target_fraction']:.4f}, "
            f"lower={hydro_soc_target['lower_fraction']:.4f}, "
            f"upper={hydro_soc_target['upper_fraction']:.4f}",
            flush=True,
        )

    def add_rolling_terminal_storage_constraints(n, snapshots):
        # Rolling horizon: restricciones terminales manuales por ventana.
        if hydro_soc_target is not None:
            add_hydro_terminal_band_constraint(
                n,
                snapshots,
                lower_fraction=hydro_soc_target["lower_fraction"],
                upper_fraction=hydro_soc_target["upper_fraction"],
                constraint_name=f"{window_label}_hydro_terminal_soc_band",
            )

        add_phs_terminal_soc_constraint(
            n,
            snapshots,
            battery_specs=battery_specs,
            constraint_name=f"{window_label}_phs_terminal_soc",
        )

        add_batterystore_min_final_soc_constraint(
            n,
            snapshots,
            min_final_soc_fraction=batterystore_min_final_soc_fraction,
            constraint_name=f"{window_label}_batterystore_min_final_soc",
        )

        add_batterystore_residual_value_to_objective(
            n,
            snapshots,
            residual_value_eur_per_mwh=batterystore_residual_value,
        )

    status, condition, hydro_soc_mode = solve_opf(
        grid,
        solver_name=solver_options["solver_name"],
        battery_specs=battery_specs,
        final_hydro_soc_fraction=None,
        initial_hydro_soc_fraction=None,
        mip_rel_gap=solver_options["mip_rel_gap"],
        time_limit=solver_options["time_limit"],
        threads=solver_options["threads"],
        mip_focus=solver_options["mip_focus"],
        method=solver_options["method"],
        crossover=solver_options["crossover"],
        numeric_focus=solver_options["numeric_focus"],
        bar_conv_tol=solver_options["bar_conv_tol"],
        feasibility_tol=solver_options["feasibility_tol"],
        optimality_tol=solver_options["optimality_tol"],
        bar_homogeneous=solver_options["bar_homogeneous"],
        line_flow_penalty=solver_options["line_flow_penalty"],
        use_line_length_scaling=solver_options["use_line_length_scaling"],
        additional_extra_functionality=add_rolling_terminal_storage_constraints,
    )

    if status != "ok" or condition not in ["optimal", "suboptimal"]:
        raise RuntimeError(
            f"{window_label} no generó solución válida: status={status}, condition={condition}"
        )

    batterystore_count, final_batterystore_soc_mwh = get_final_batterystore_soc_mwh(grid)
    batterystore_residual_value_deducted = (
        batterystore_residual_value * final_batterystore_soc_mwh
    )
    objective_with_residual_value = float(grid.objective)
    total_cost_without_residual_value = (
        objective_with_residual_value + batterystore_residual_value_deducted
    )

    print(
        f"{window_label}: BatteryStore residual value summary | "
        f"n_BatteryStore={batterystore_count}, "
        f"residual_value={batterystore_residual_value:.2f} €/MWh, "
        f"final_SOC={final_batterystore_soc_mwh:.4f} MWh, "
        f"deducted={batterystore_residual_value_deducted:.2f} €, "
        f"solver_objective={objective_with_residual_value:.2f} €, "
        f"cost_without_residual={total_cost_without_residual_value:.2f} €",
        flush=True,
    )

    if batterystore_min_final_soc_percent > 0.0:
        print(
            f"{window_label}: BatteryStore minimum final SOC summary | "
            f"min_final_soc={batterystore_min_final_soc_percent:.2f} %, "
            f"n_BatteryStore={batterystore_min_final_soc_count}, "
            f"min_final_SOC_total={batterystore_min_final_soc_mwh_total:.4f} MWh, "
            f"final_SOC={final_batterystore_soc_mwh:.4f} MWh",
            flush=True,
        )

    solver_result_values = {
        "objective": objective_with_residual_value,
        "status": status,
        "condition": condition,
        "hydro_soc_mode": hydro_soc_mode,
        "window_label": window_label,
        "rolling_batterystore_residual_value_eur_per_mwh": batterystore_residual_value,
        "batterystore_count_for_residual_value": batterystore_count,
        "final_BatteryStore_SOC_MWh": final_batterystore_soc_mwh,
        "BatteryStore_min_final_SOC_percent": batterystore_min_final_soc_percent,
        "BatteryStore_min_final_SOC_MWh_total": batterystore_min_final_soc_mwh_total,
        "BatteryStore_residual_value_deducted_eur": batterystore_residual_value_deducted,
        "objective_with_batterystore_residual_value": objective_with_residual_value,
        "total_cost_without_batterystore_residual_value": total_cost_without_residual_value,
    }

    if hydro_soc_target is not None:
        solver_result_values.update({
            "rolling_hydro_target_fraction": hydro_soc_target["target_fraction"],
            "rolling_hydro_lower_fraction": hydro_soc_target["lower_fraction"],
            "rolling_hydro_upper_fraction": hydro_soc_target["upper_fraction"],
        })

    solver_results = pd.DataFrame({"value": solver_result_values})

    df_available_renewable = _build_available_renewable(
        df_solar_node_profiles,
        df_wind_node_profiles,
    )

    # Rolling horizon: por ventana solo extraemos tablas. No se generan
    # gráficos, no se crea Sankey y no se escribe Excel completo.
    result_tables = extract_multiperiod_result_tables(
        grid,
        df_available_renewable,
        include_renewable_detail=False,
        include_battery_kpis=False,
    )

    return {
        "window_label": window_label,
        "start": window_snapshots[0],
        "end": window_snapshots[-1],
        "hours": len(window_snapshots),
        "grid": grid,
        "solver_results": solver_results,
        "status": status,
        "condition": condition,
        "hydro_soc_mode": hydro_soc_mode,
        "hydro_soc_target": hydro_soc_target,
        "objective": objective_with_residual_value,
        "batterystore_residual_kpis": {
            "rolling_batterystore_residual_value_eur_per_mwh": batterystore_residual_value,
            "batterystore_count_for_residual_value": batterystore_count,
            "final_BatteryStore_SOC_MWh": final_batterystore_soc_mwh,
            "BatteryStore_min_final_SOC_percent": batterystore_min_final_soc_percent,
            "BatteryStore_min_final_SOC_MWh_total": batterystore_min_final_soc_mwh_total,
            "BatteryStore_residual_value_deducted_eur": batterystore_residual_value_deducted,
            "objective_with_batterystore_residual_value": objective_with_residual_value,
            "total_cost_without_batterystore_residual_value": total_cost_without_residual_value,
        },
        "output_file": None,
        "tables": result_tables,
    }


def aggregate_rolling_results(window_results: list[dict]) -> dict:
    hourly_keys = [
        "dispatch",
        "storage_dispatch",
        "storage_soc",
        "line_flows",
        "prices",
        "loads",
        "ccgt_marginal_cost",
        "available_renewable",
        "generators_p",
        "generators_marginal_cost",
        "links_p0",
        "links_p1",
        "loads_p",
        "stores_e",
        "lines_p0",
        "buses_marginal_price",
    ]

    aggregated = {}

    for key in hourly_keys:
        pieces = []
        for result in window_results:
            tables = result.get("tables", {})
            if key in tables and tables[key] is not None:
                pieces.append(
                    _normalize_temporal_result_table(
                        tables[key],
                        key,
                        result,
                    )
                )

        if not pieces:
            continue

        combined = pd.concat(pieces, axis=0)

        # Si en el futuro se usan ventanas solapadas, aquí debe decidirse
        # qué parte de la ventana se conserva antes de concatenar.
        duplicated = combined.index.duplicated(keep="first")
        if duplicated.any():
            duplicated_snapshots = combined.index[duplicated]
            first_duplicates = list(duplicated_snapshots[:10])
            warnings.warn(
                f"Se han eliminado {duplicated.sum()} snapshots duplicados en {key}. "
                f"Primeros duplicados: {first_duplicates}"
            )
            combined = combined.loc[~duplicated]

        aggregated[key] = combined.sort_index()

    return aggregated


def _build_window_kpis(window_results: list[dict]) -> pd.DataFrame:
    rows = []
    for result in window_results:
        row = {
            "window": result["window_label"],
            "start": result["start"],
            "end": result["end"],
            "hours": result["hours"],
            "status": result["status"],
            "condition": result["condition"],
            "hydro_soc_mode": result["hydro_soc_mode"],
            "objective": result["objective"],
            "output_file": result["output_file"],
        }

        hydro_soc_target = result.get("hydro_soc_target")
        if hydro_soc_target is not None:
            row.update({
                "hydro_target_window_end": hydro_soc_target["window_end"],
                "hydro_target_fraction": hydro_soc_target["target_fraction"],
                "hydro_lower_fraction": hydro_soc_target["lower_fraction"],
                "hydro_upper_fraction": hydro_soc_target["upper_fraction"],
            })

        batterystore_residual_kpis = result.get("batterystore_residual_kpis")
        if batterystore_residual_kpis is not None:
            row.update(batterystore_residual_kpis)

        rows.append(row)

    return pd.DataFrame(rows)


def _build_aggregated_kpis(
    aggregated_results: dict,
    window_results: list[dict],
) -> pd.DataFrame:
    rows = [
        {"metric": "windows_completed", "value": len(window_results)},
        {
            "metric": "objective_sum",
            "value": sum(float(result["objective"]) for result in window_results),
        },
    ]

    objective_with_residual = sum(
        float(
            result.get("batterystore_residual_kpis", {}).get(
                "objective_with_batterystore_residual_value",
                result["objective"],
            )
        )
        for result in window_results
    )
    residual_value_deducted = sum(
        float(
            result.get("batterystore_residual_kpis", {}).get(
                "BatteryStore_residual_value_deducted_eur",
                0.0,
            )
        )
        for result in window_results
    )
    total_cost_without_residual = sum(
        float(
            result.get("batterystore_residual_kpis", {}).get(
                "total_cost_without_batterystore_residual_value",
                result["objective"],
            )
        )
        for result in window_results
    )

    rows.extend([
        {
            "metric": "Objective with BatteryStore residual value (€)",
            "value": objective_with_residual,
        },
        {
            "metric": "BatteryStore residual value deducted (€)",
            "value": residual_value_deducted,
        },
        {
            "metric": "Total cost without BatteryStore residual value (€)",
            "value": total_cost_without_residual,
        },
    ])

    if "loads" in aggregated_results:
        loads = aggregated_results["loads"]
        if "Total_load" in loads.columns:
            rows.append({
                "metric": "total_load_mwh",
                "value": float(loads["Total_load"].sum()),
            })

    if "dispatch" in aggregated_results:
        dispatch = aggregated_results["dispatch"]
        for col in dispatch.columns:
            rows.append({
                "metric": f"dispatch_{col}_mwh",
                "value": float(dispatch[col].sum()),
            })

    if "renewable_detail" in aggregated_results:
        renewable = aggregated_results["renewable_detail"]
        for col in ["Total_available", "Total_used", "Total_curtailment"]:
            if col in renewable.columns:
                rows.append({
                    "metric": f"renewable_{col}_mwh",
                    "value": float(renewable[col].sum()),
                })

    return pd.DataFrame(rows)


def _build_aggregate_grid_for_export(
    window_results: list[dict],
    aggregated_results: dict,
):
    if not window_results:
        raise ValueError("No hay resultados de ventanas para construir la red agregada.")

    template_grid = window_results[0]["grid"]

    if "generators_p" in aggregated_results:
        snapshots = pd.DatetimeIndex(aggregated_results["generators_p"].index)
    elif "dispatch" in aggregated_results:
        snapshots = pd.DatetimeIndex(aggregated_results["dispatch"].index)
    else:
        raise ValueError("No hay resultados horarios para construir la red agregada.")

    objective = sum(float(result["objective"]) for result in window_results)

    aggregate_grid = SimpleNamespace()
    aggregate_grid.snapshots = snapshots
    aggregate_grid.objective = objective
    aggregate_grid.snapshot_weightings = pd.DataFrame(
        {"objective": 1.0},
        index=snapshots,
    )

    # Tablas estáticas: las capacidades son fijas en rolling horizon.
    aggregate_grid.generators = template_grid.generators.copy()
    aggregate_grid.links = template_grid.links.copy()
    aggregate_grid.loads = template_grid.loads.copy()
    aggregate_grid.stores = template_grid.stores.copy()
    aggregate_grid.lines = template_grid.lines.copy()
    aggregate_grid.buses = template_grid.buses.copy()
    aggregate_grid.storage_units = template_grid.storage_units.copy()

    aggregate_grid.generators_t = SimpleNamespace(
        p=aggregated_results.get("generators_p", pd.DataFrame(index=snapshots)),
        marginal_cost=aggregated_results.get(
            "generators_marginal_cost",
            pd.DataFrame(index=snapshots),
        ),
    )
    aggregate_grid.links_t = SimpleNamespace(
        p0=aggregated_results.get("links_p0", pd.DataFrame(index=snapshots)),
        p1=aggregated_results.get("links_p1", pd.DataFrame(index=snapshots)),
    )
    aggregate_grid.loads_t = SimpleNamespace(
        p=aggregated_results.get("loads_p", pd.DataFrame(index=snapshots)),
    )
    aggregate_grid.stores_t = SimpleNamespace(
        e=aggregated_results.get("stores_e", pd.DataFrame(index=snapshots)),
    )
    aggregate_grid.lines_t = SimpleNamespace(
        p0=aggregated_results.get("lines_p0", pd.DataFrame(index=snapshots)),
    )
    aggregate_grid.buses_t = SimpleNamespace(
        marginal_price=aggregated_results.get(
            "buses_marginal_price",
            pd.DataFrame(index=snapshots),
        ),
    )

    return aggregate_grid


def save_rolling_horizon_results(
    config: dict,
    window_results: list[dict],
    aggregated_results: dict,
    errors: list[dict] | None = None,
    partial: bool = False,
) -> Path:
    prepared = config["prepared_inputs"]
    run_timestamp = config["run_timestamp"]
    suffix = "_partial" if partial else ""
    output_file = Path(f"results_rolling_horizon{suffix}_{run_timestamp}.xlsx")

    settings = prepared["df_SYS_settings_full"]["SYSTEM PARAMETERS"].copy()
    settings["rolling_horizon_enabled"] = True
    settings["rolling_horizon_days"] = config["rolling_horizon_days"]
    settings["rolling_hydro_soc_band_percent"] = config.get(
        "rolling_hydro_soc_band_percent",
        5.0,
    )
    settings["rolling_batterystore_residual_value_eur_per_mwh"] = config.get(
        "rolling_batterystore_residual_value_eur_per_mwh",
        0.0,
    )
    settings["rolling_batterystore_min_final_soc_percent"] = config.get(
        "rolling_batterystore_min_final_soc_percent",
        0.0,
    )
    settings["rolling_windows_completed"] = len(window_results)
    settings["rolling_partial_results"] = partial
    df_settings = pd.DataFrame({"ROLLING HORIZON SETTINGS": settings})

    df_window_kpis = _build_window_kpis(window_results)
    df_aggregated_kpis = _build_aggregated_kpis(aggregated_results, window_results)

    if not partial and window_results:
        aggregate_grid = _build_aggregate_grid_for_export(
            window_results,
            aggregated_results,
        )

        if "available_renewable" not in aggregated_results:
            raise ValueError(
                "No se han acumulado disponibilidades renovables para el export final."
            )

        solver_results = pd.DataFrame({
            "value": {
                "objective": aggregate_grid.objective,
                "status": "ok",
                "condition": "rolling_horizon_completed",
                "windows_completed": len(window_results),
                "objective_with_batterystore_residual_value": sum(
                    float(
                        result.get("batterystore_residual_kpis", {}).get(
                            "objective_with_batterystore_residual_value",
                            result["objective"],
                        )
                    )
                    for result in window_results
                ),
                "BatteryStore_residual_value_deducted_eur": sum(
                    float(
                        result.get("batterystore_residual_kpis", {}).get(
                            "BatteryStore_residual_value_deducted_eur",
                            0.0,
                        )
                    )
                    for result in window_results
                ),
                "total_cost_without_batterystore_residual_value": sum(
                    float(
                        result.get("batterystore_residual_kpis", {}).get(
                            "total_cost_without_batterystore_residual_value",
                            result["objective"],
                        )
                    )
                    for result in window_results
                ),
            }
        })

        # Rolling horizon: el export completo, gráficos y Sankey se ejecutan
        # una sola vez al final, con todos los snapshots concatenados.
        export_multiperiod_results(
            aggregate_grid,
            prepared["df_SYS_settings_full"],
            aggregated_results["available_renewable"],
            prepared["CFsolar"],
            prepared["CFwind"],
            prepared["df_StorageUnit"],
            solver_results,
            output_file=output_file,
            print_index_diagnostics=True,
        )

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            df_settings.to_excel(writer, sheet_name="Rolling settings")
            df_window_kpis.to_excel(writer, sheet_name="Window KPIs", index=False)
            df_aggregated_kpis.to_excel(writer, sheet_name="Aggregated KPIs", index=False)

            if errors:
                pd.DataFrame(errors).to_excel(writer, sheet_name="Errors", index=False)

        return output_file

    sheet_names = {
        "dispatch": "dispatch",
        "renewable_detail": "Solar and wind",
        "storage_dispatch": "storage dispatch",
        "storage_soc": "storage SOC",
        "line_flows": "line flows",
        "prices": "prices",
        "loads": "loads",
        "ccgt_marginal_cost": "CCGT marginal cost",
    }

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_settings.to_excel(writer, sheet_name="Settings")
        df_window_kpis.to_excel(writer, sheet_name="Window KPIs", index=False)
        df_aggregated_kpis.to_excel(writer, sheet_name="Aggregated KPIs", index=False)

        for key, sheet_name in sheet_names.items():
            if key in aggregated_results:
                aggregated_results[key].to_excel(writer, sheet_name=sheet_name)

        if errors:
            pd.DataFrame(errors).to_excel(writer, sheet_name="Errors", index=False)

    return output_file


def run_rolling_horizon(config: dict) -> Path:
    input_path = Path(config["input_path"])
    data = config["data"]
    system_parameters = config["system_parameters"]
    progress_callback = config.get("progress_callback")

    def report_progress(value: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(value, message)

    rolling_horizon_days = validate_rolling_horizon_days(
        system_parameters.get("rolling_horizon_days", 3)
    )
    rolling_hydro_soc_band_percent = validate_rolling_hydro_soc_band_percent(
        system_parameters.get("rolling_hydro_soc_band_percent", 5.0)
    )
    rolling_hydro_soc_band = rolling_hydro_soc_band_percent / 100.0
    batterystore_residual_value = validate_rolling_batterystore_residual_value(
        system_parameters.get(
            "rolling_batterystore_residual_value_eur_per_mwh",
            0.0,
        )
    )
    batterystore_min_final_soc_percent = (
        validate_rolling_batterystore_min_final_soc_percent(
            system_parameters.get(
                "rolling_batterystore_min_final_soc_percent",
                0.0,
            )
        )
    )

    report_progress(5, "rolling horizon: preparando datos largos")
    df_SYS_settings = build_sys_settings_from_gui(system_parameters)

    prepared_inputs = _prepare_rolling_horizon_inputs(
        input_path=input_path,
        data=data,
        system_parameters=system_parameters,
        df_SYS_settings=df_SYS_settings,
    )

    full_snapshots = prepared_inputs["full_snapshots"]
    windows = generate_rolling_windows(full_snapshots, rolling_horizon_days)

    if not windows:
        raise ValueError("No se ha generado ninguna ventana rolling horizon.")

    hydro_soc_target_trajectory = build_hydro_soc_target_trajectory(
        startdate=prepared_inputs["startdate"],
        fecha_final=prepared_inputs["fecha_final"],
        rolling_windows=windows,
        df_embalses_original=prepared_inputs["df_embalses_original"],
        hydro_band_fraction=rolling_hydro_soc_band,
    )

    run_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    rolling_config = {
        "prepared_inputs": prepared_inputs,
        "rolling_horizon_days": rolling_horizon_days,
        "rolling_hydro_soc_band_percent": rolling_hydro_soc_band_percent,
        "rolling_hydro_soc_band": rolling_hydro_soc_band,
        "rolling_batterystore_residual_value_eur_per_mwh": batterystore_residual_value,
        "rolling_batterystore_min_final_soc_percent": batterystore_min_final_soc_percent,
        "hydro_soc_target_trajectory": hydro_soc_target_trajectory,
        "run_timestamp": run_timestamp,
    }

    window_results = []
    errors = []
    previous_storage_state = None
    previous_generator_power = None
    total_windows = len(windows)

    for window_number, window_snapshots in enumerate(windows, start=1):
        window_label = f"rolling_window_{window_number:03d}"
        window_start = window_snapshots[0]
        window_end = window_snapshots[-1]
        hydro_soc_target = hydro_soc_target_trajectory.get(window_label)

        if hydro_soc_target is None:
            raise ValueError(f"No hay trayectoria hydro para {window_label}.")

        progress_value = 10 + int((window_number - 1) / total_windows * 80)
        report_progress(
            progress_value,
            (
                f"{window_label}/{total_windows}: "
                f"{window_start:%Y-%m-%d %H:%M} - {window_end:%Y-%m-%d %H:%M}"
            ),
        )

        window_config = dict(rolling_config)
        window_config.update({
            "window_label": window_label,
            "window_number": window_number,
            "total_windows": total_windows,
            "hydro_soc_target": hydro_soc_target,
        })

        print(
            f"{window_label}: fecha final de ventana {hydro_soc_target['window_end']}; "
            f"target hydro SOC={hydro_soc_target['target_fraction']:.4f}; "
            f"lower={hydro_soc_target['lower_fraction']:.4f}; "
            f"upper={hydro_soc_target['upper_fraction']:.4f}",
            flush=True,
        )

        try:
            result = run_single_rolling_window(
                window_config,
                window_snapshots,
                previous_storage_state=previous_storage_state,
                previous_generator_power=previous_generator_power,
            )
        except Exception as exc:
            error_text = traceback.format_exc()
            errors.append({
                "window": window_label,
                "start": window_start,
                "end": window_end,
                "error": str(exc),
                "traceback": error_text,
            })

            aggregated_partial = aggregate_rolling_results(window_results)
            partial_file = save_rolling_horizon_results(
                rolling_config,
                window_results,
                aggregated_partial,
                errors=errors,
                partial=True,
            )

            raise RuntimeError(
                f"Rolling horizon detenido en {window_label}. "
                f"Resultados parciales guardados en {partial_file}.\n{error_text}"
            ) from exc

        window_results.append(result)
        previous_storage_state = extract_final_storage_state(result)
        previous_generator_power = extract_final_generator_power(result)

        report_progress(
            10 + int(window_number / total_windows * 80),
            (
                f"Ventana {window_number}/{total_windows} resuelta. "
                "Resultados añadidos al acumulado."
            ),
        )

    report_progress(95, "rolling horizon: agregando resultados")
    aggregated_results = aggregate_rolling_results(window_results)
    output_file = save_rolling_horizon_results(
        rolling_config,
        window_results,
        aggregated_results,
        errors=errors,
        partial=False,
    )

    report_progress(100, "rolling horizon: resultados exportados")
    return output_file


def run_program(
    input_file: str | Path | None = None,
    system_parameters: dict | None = None,
    progress_callback=None
) -> Path:
    
    logging.getLogger("pypsa").setLevel(logging.ERROR)
    logging.getLogger("linopy").setLevel(logging.ERROR)

    warnings.filterwarnings("ignore", category=pd.errors.PerformanceWarning)
    warnings.filterwarnings("ignore", message="Tight layout not applied*")

    def report_progress(value: int, message: str) -> None:
        if progress_callback is not None:
            progress_callback(value, message)

    report_progress(0, "iniciando")
   
    if input_file is None:
        input_path = get_default_input_file()
    else:
        input_path = Path(input_file).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {input_path}")
    

    data = leerhojas(input_path)
    report_progress(10, "datos de entrada leídos")
    
    if system_parameters is None:
        raise ValueError("No se han recibido los parámetros del sistema desde la GUI.")

    # Rolling horizon: si está desactivado, se continúa con el flujo antiguo.
    rolling_horizon_enabled = parse_bool_setting(
        system_parameters.get("rolling_horizon_enabled", False)
    )

    if rolling_horizon_enabled:
        if parse_bool_setting(
            system_parameters.get("intermediate_storage_constraints_enabled", False)
        ):
            print(
                "Rolling horizon activado: se ignoran las restricciones "
                "intermedias hydro/PHS porque solo aplican sin rolling horizon.",
                flush=True,
            )
        validate_rolling_horizon_days(system_parameters.get("rolling_horizon_days", 3))
        validate_rolling_hydro_soc_band_percent(
            system_parameters.get("rolling_hydro_soc_band_percent", 5.0)
        )
        validate_rolling_batterystore_residual_value(
            system_parameters.get(
                "rolling_batterystore_residual_value_eur_per_mwh",
                0.0,
            )
        )
        validate_rolling_batterystore_min_final_soc_percent(
            system_parameters.get(
                "rolling_batterystore_min_final_soc_percent",
                0.0,
            )
        )
        return run_rolling_horizon(
            {
                "input_path": input_path,
                "data": data,
                "system_parameters": system_parameters,
                "progress_callback": progress_callback,
            }
        )

    df_SYS_settings = build_sys_settings_from_gui(system_parameters)
    report_progress(20, "parámetros de la GUI cargados")
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    startdate = params["Start date (dd/mm/aaaa)"]
    duration = params["Simulation duration (days)"]
    intermediate_storage_constraints_enabled = (
        horizon == "Multiperiod"
        and parse_bool_setting(
            system_parameters.get("intermediate_storage_constraints_enabled", False)
        )
    )

    if intermediate_storage_constraints_enabled:
        intermediate_storage_constraint_days = validate_intermediate_storage_constraint_days(
            system_parameters.get("intermediate_storage_constraint_days", 3)
        )
        intermediate_hydro_soc_band_percent = validate_intermediate_hydro_soc_band_percent(
            system_parameters.get("intermediate_hydro_soc_band_percent", 0.5)
        )
    else:
        intermediate_storage_constraint_days = None
        intermediate_hydro_soc_band_percent = None


    if horizon == "Static":
        static_snapshot_datetime = params["Static snapshot datetime"]
        startdate = static_snapshot_datetime.date()
        duration = 1
        intermediate_storage_constraints_enabled = False
        
    

    # -- TRATAMIENTO DE PERFILES DE DEMANDA de ESPAÑA -- 

    BASE_DIR = Path(__file__).resolve().parent
    rutaDemandaXccaa = BASE_DIR / "System_data" / "demanda_anual_ccaa_REE_2015_2025.xlsx"
    df_consumos_anuales_CCAA = pd.read_excel(rutaDemandaXccaa, sheet_name="demanda_ccaa_largo", usecols="A,B,E")
    df_consumos_anuales_CCAA = df_consumos_anuales_CCAA.dropna()

    demr_folder = BASE_DIR / "System_data" / "DemandaDiariaSistemaElectricoPeninsular"
    df_demanda_ccaa = build_hourly_demand_by_region(
        df_consumos_anuales_CCAA=df_consumos_anuales_CCAA,
        demr_folder=demr_folder,
        startdate=startdate,
        days=duration,
    )

    ruta2 = BASE_DIR / "System_data" / "2013PyPSA_Network.xlsx"

    df_PyPSA2013_load_profiles = pd.read_excel(
        ruta2,
        sheet_name="loads_timeseries"
    )

    df_monthly_node_weights_ESP = build_monthly_nodal_load_weights_ES(
    df_pypsa_load_profiles=df_PyPSA2013_load_profiles,
    node_to_region=node_to_region,
    exclude_portugal=True
    )

    df_demanda_nodal_ESP = build_hourly_nodal_demand(
    df_demanda_ccaa=df_demanda_ccaa,
    df_monthly_node_weights=df_monthly_node_weights_ESP
    )
    df_demanda_nodal_ESP = df_demanda_nodal_ESP.set_index("time")

    # -- TRATAMIENTO DE PERFILES DE DEMANDA de PORTUGAL -- 

    ruta_demanda_PT = BASE_DIR / "System_data" / "DemandaPT.xlsx"

    df_total_demand_pt = pd.read_excel(ruta_demanda_PT, sheet_name="Portugal demand TS", usecols="A:B")
    df_region_demand = pd.read_excel(ruta_demanda_PT, sheet_name="Consumos mensuales por región", usecols="A:F", header=1)

    df_demand_regions_hourly = regional_hourly_demand_builder(
    df_regional_demand=df_region_demand,
    df_total_hourly_demand=df_total_demand_pt,
    startdate=startdate,
    days=duration)

    df_monthly_node_weights_PT = build_monthly_nodal_load_weights_PT(
    df_pypsa_load_profiles=df_PyPSA2013_load_profiles,
    node_to_region=node_to_region_PT,
    exclude_spain=True)

    df_demanda_nodal_PT = build_hourly_nodal_demand_PT(
        df_demand_PT_regions=df_demand_regions_hourly,
        df_monthly_node_weights=df_monthly_node_weights_PT)
    df_demanda_nodal_PT = df_demanda_nodal_PT.set_index("time")
  
    # Fusionamos los dataframe de series de demanda de PT y ESP
    df_demanda_nodal = pd.concat([df_demanda_nodal_PT, df_demanda_nodal_ESP], axis=1)
    df_demanda_nodal.index = pd.to_datetime(df_demanda_nodal.index)

    # -- TRATAMIENTO DE PERFILES DE RENOVABLES --
    ruta = BASE_DIR / "GridInputs.xlsx"
    # Lectura de potencia instalada
    df_solar_instaled_capacity = pd.read_excel(ruta, engine="openpyxl", sheet_name="Gen_PV_and_Wind", usecols="G:AB", skiprows=2, nrows=14 )
    df_wind_instaled_capacity = pd.read_excel(ruta, engine="openpyxl", sheet_name="Gen_PV_and_Wind", usecols="G:AB", skiprows=19, nrows=14 )
    df_wind_instaled_capacity.columns = (
        df_wind_instaled_capacity.columns
        .str.replace(r"\.\d+$", "", regex=True)
        .str.strip()
    )
    df_solar_instaled_capacity.columns = (
        df_solar_instaled_capacity.columns
        .str.replace(r"\.\d+$", "", regex=True)
        .str.strip()
    )

    # Lectura de los perfiles por unidad de renewables ninja
    rutaTSRenewablesninja = BASE_DIR / "System_data" / "RenewablesNinja_Time_Series.xlsx"
    df_solar_profiles = pd.read_excel(rutaTSRenewablesninja, sheet_name="TS_PV_Profiles", usecols="A:W")
    df_wind_profiles = pd.read_excel(rutaTSRenewablesninja, sheet_name="TS_Wind_Profiles", usecols="A:W")
    CFwind = df_wind_profiles.drop("time", axis=1).mean().mean() #Factor de capacidad de la generación eólica
    CFsolar = df_solar_profiles.drop("time", axis=1).mean().mean() #Factor de capacidad de la generación FV
  
    # Lectura de los pesos nodales
    df_node_weights = pd.read_excel(ruta, sheet_name="Gen_PV_and_Wind", usecols="B:E", skiprows=2, nrows=97 )
    df_solar_node_weights = df_node_weights[df_node_weights["Renewable Type"]=="PV"]
    df_wind_node_weights = df_node_weights[df_node_weights["Renewable Type"]=="Wind"]

    # Llamamos a la función para construir las series horarias por nodo de FV
    df_solar_node_profiles = renewable_profile_builder(
    df_installed_capacity=df_solar_instaled_capacity,
    df_node_weights=df_solar_node_weights,
    df_profiles=df_solar_profiles,
    days=duration,
    startdate=startdate
    )

    # Llamamos a la función para construir las series horarias por nodo de eólica
    df_wind_node_profiles = renewable_profile_builder(
    df_installed_capacity=df_wind_instaled_capacity,
    df_node_weights=df_wind_node_weights,
    df_profiles=df_wind_profiles,
    days=duration,
    startdate=startdate,
    )
    # -- LECTURA DE PRECIOS DE FRANCIA --
    rutaFrancia = BASE_DIR / "System_data" / "precios_francia_2015_2024.xlsx"
    df_TS_Energy_Prices = pd.read_excel(rutaFrancia, sheet_name="PreciosFrancia", usecols="A:B")
    df_TS_Energy_Prices["time"] = pd.to_datetime(df_TS_Energy_Prices["time"]).dt.round("h")
    df_TS_Energy_Prices = df_TS_Energy_Prices.set_index("time")
    df_TS_Energy_Prices["Precio Francia (€/MWh)"] = df_TS_Energy_Prices["Precio Francia (€/MWh)"].clip(lower=0)


    # -- CALCULO DE LAS SERIES TEMPORALES DE INFLOW PARA GENERADORES HIDRÁULICOS SIN BOMBEO ESCALADO A PARTIR DEL RUNOFF --
    df_hydro_inflow_scaled = build_hydro_inflow(
    base_dir=BASE_DIR,
    startdate=startdate,
    days=duration)

    """def hydro_inflow_annual_from_scaled_df(
        df_hydro_inflow_scaled,
        year=None,
        value_name="Annual hydro inflow (MWh)"
    ):
        df = df_hydro_inflow_scaled.copy()

        # Asegurar índice temporal
        df.index = pd.to_datetime(df.index)

        # Quedarse solo con columnas numéricas
        hydro_cols = df.select_dtypes(include="number").columns

        if len(hydro_cols) == 0:
            raise ValueError("No hay columnas numéricas de inflow hidroeléctrico.")

        # Inflow total anual del sistema
        annual_hydro_inflow = df[hydro_cols].sum().sum()

        # Detectar año si no se pasa explícitamente
        if year is None:
            years = df.index.year.unique()

            if len(years) != 1:
                raise ValueError(
                    f"El DataFrame contiene varios años: {list(years)}. "
                    "Pasa el argumento year explícitamente."
                )

            year = int(years[0])

        return pd.DataFrame({
            "year": [year],
            value_name: [annual_hydro_inflow]
        })

    df_annual_hydro = hydro_inflow_annual_from_scaled_df(df_hydro_inflow_scaled)
    print(df_annual_hydro)"""

    # -- SERIE TEMPORAL DE P_MAX_PU DE GENERADORES ROR ESCALADA A PARTIR DEL RUNOFF --
    df_ror_p_max_pu_scaled = build_ror_p_max_pu(base_dir=BASE_DIR, startdate=startdate, days=duration)

    # -- CALCULO DEL PORCENTAJE INICIAL DE LLENADO DE LOS EMBALSES HYDRO
    BASE_DIR = Path(__file__).resolve().parent

    ruta_llenado_embalses = BASE_DIR / "System_data" / "Embalses.xlsx"

    df_embalses_original = pd.read_excel(ruta_llenado_embalses, usecols="D:G")

    df_embalses_original = df_embalses_original[
        df_embalses_original["ELECTRICO_FLAG"] == 1
    ]

    df_embalses_initial_SOC = get_embalses_closest_date(
        df=df_embalses_original,
        target_date=startdate
    )

    initial_soc_fraction = (
        df_embalses_initial_SOC["AGUA_ACTUAL"].sum()
        / df_embalses_initial_SOC["AGUA_TOTAL"].sum()
    )

    fecha_inicio = pd.to_datetime(startdate, dayfirst=True)
    fecha_final = fecha_inicio + pd.to_timedelta(duration, unit="D")

    df_embalses_final_SOC = get_embalses_closest_date(
        df=df_embalses_original,
        target_date=fecha_final
    )

    final_soc_fraction = (
        df_embalses_final_SOC["AGUA_ACTUAL"].sum()
        / df_embalses_final_SOC["AGUA_TOTAL"].sum()
    )

    # -- CÁLCULO DEL CAPITAL RECOVERY FACTOR --

    economic_settings = build_battery_economic_settings_from_gui(system_parameters)
    
    if horizon=="Multiperiod" and economic_settings.iloc[0, 0] is not None:
        def crf(i, n):
            if i == 0:
                return 1 / n
            return (i * (1 + i)**n) / ((1 + i)**n - 1)
        
        interest = economic_settings.iloc[0, 0]/100
        lifetime = economic_settings.iloc[1, 0]
        CRF = crf(interest, lifetime)
    else:
        CRF = 0


    df_Net_Buses = data["Net_Buses"]
    df_Net_Lines = data["Net_Lines"]
    df_Gen_Dispatchable = data["Gen_Dispatchable"]
    df_StorageUnit = data["StorageUnit"]
    df_Grid_connection = data["Grid_connection"]

    grid = build_network(df_SYS_settings)
    report_progress(30, "red creada")

    # -- OBTENEMOS DATAFRAME PARA EL CÁLCULO DEL PRECIO MARGINAL DEL CICLO COMBINADO --
    gas_price, co2_price = CCGT_dataframe_treatment(BASE_DIR, startdate, duration)
    gas_price = daily_to_snapshots(gas_price, grid.snapshots)
    co2_price = daily_to_snapshots(co2_price, grid.snapshots)


    add_buses(grid, df_Net_Buses)
    report_progress(38, "buses añadidos")


    add_lines(grid, df_Net_Lines)
    report_progress(45, "líneas añadidas")
    add_loads(grid, df_demanda_nodal, df_SYS_settings)
    grid.loads_t.p_set.sum(axis=1).plot()
    annual_load_twh = grid.loads_t.p_set.sum().sum() / 1e6

    print("Carga total (TWh): ", annual_load_twh)

    report_progress(52, "cargas añadidas")
    add_dispatchable_generators(grid, df_Gen_Dispatchable, gas_price, co2_price, df_ror_p_max_pu_scaled)
    report_progress(58, "generadores despachables añadidos")
    add_renewable_generator(grid, params, df_solar_node_profiles, df_wind_node_profiles)

    report_progress(64, "generadores renovables añadidos")
    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)
    report_progress(70, "conexión a red añadida")


    """if grid.buses.x.isna().any() or grid.buses.y.isna().any():
        print("There are buses for which latitude/longitude were not specified, therefore the grid will not be drawn")
    else:
        drawrealgrid(grid, df_Net_Buses, "Rediberica.png")"""


    if horizon == "Multiperiod" and pd.notna(df_StorageUnit.loc[0, "LOCATION"]):
        battery_specs = add_storage_as_store_links(df_SYS_settings, grid, df_StorageUnit, CRF, df_hydro_inflow_scaled, initial_soc_fraction)
    elif horizon == "Static":
        battery_specs=None
    else:
        battery_specs=None
    report_progress(76, "almacenamiento configurado")
    
    
    hydro_units = grid.storage_units.index[
    grid.storage_units.carrier == "hydro"
    ]
    grid.storage_units_t.inflow.loc[:, hydro_units] = df_hydro_inflow_scaled[hydro_units]


    intermediate_extra_functionality = None
    if intermediate_storage_constraints_enabled:
        intermediate_block_windows = generate_rolling_windows(
            grid.snapshots,
            intermediate_storage_constraint_days,
        )
        intermediate_hydro_soc_band_fraction = intermediate_hydro_soc_band_percent / 100.0
        intermediate_hydro_soc_target_trajectory = build_hydro_soc_target_trajectory(
            startdate=fecha_inicio,
            fecha_final=fecha_final,
            rolling_windows=intermediate_block_windows,
            df_embalses_original=df_embalses_original,
            hydro_band_fraction=intermediate_hydro_soc_band_fraction,
            log_prefix="Intermediate storage constraints",
        )

        print(
            "Intermediate storage constraints enabled: "
            f"n_blocks={len(intermediate_block_windows)}, "
            f"block_days={intermediate_storage_constraint_days}, "
            f"horizon_start={grid.snapshots[0]}, "
            f"horizon_end={grid.snapshots[-1]}, "
            f"fecha_final={fecha_final}, "
            f"hydro_band={intermediate_hydro_soc_band_fraction:.4f} "
            f"({intermediate_hydro_soc_band_fraction:.2%}).",
            flush=True,
        )

        def intermediate_extra_functionality(n, snapshots):
            # Simulación anual con información perfecta: una sola optimización,
            # pero con restricciones terminales intermedias para hydro/PHS.
            add_intermediate_storage_terminal_constraints(
                n,
                intermediate_block_windows,
                intermediate_hydro_soc_target_trajectory,
                battery_specs,
            )



    report_progress(80, "optimizando")

    def as_optional_float(value):
        if value is None:
            return None
        if isinstance(value, str) and value.strip().lower() in ["", "none", "nan"]:
            return None
        return float(value)


    def as_optional_int(value, zero_is_none=False):
        if value is None:
            return None
        if isinstance(value, str) and value.strip().lower() in ["", "none", "nan"]:
            return None

        value = int(float(value))

        if zero_is_none and value == 0:
            return None

        return value


    mip_rel_gap = as_optional_float(params.get("mip_rel_gap"))

    time_limit = as_optional_int(params.get("time_limit"), zero_is_none=True)
    threads = as_optional_int(params.get("threads"), zero_is_none=False)

   

    mip_focus = as_optional_int(params.get("mip_focus"))
    solver_name = str(params.get("solver_name"))

    method = as_optional_int(params.get("method"))
    bar_homogeneous = as_optional_int(params.get("bar_homogeneous"), zero_is_none=False)
    crossover = as_optional_int(params.get("crossover"))

    numeric_focus = as_optional_int(params.get("numeric_focus"))
    bar_conv_tol = as_optional_float(params.get("bar_conv_tol"))
    feasibility_tol = as_optional_float(params.get("feasibility_tol"))
    optimality_tol = as_optional_float(params.get("optimality_tol"))

    line_flow_penalty = as_optional_float(params.get("line_flow_penalty"))
    if line_flow_penalty is None:
        line_flow_penalty = 0.0

    use_line_length_scaling_raw = params.get("use_line_length_scaling", True)

    if isinstance(use_line_length_scaling_raw, str):
        use_line_length_scaling = use_line_length_scaling_raw.lower() in [
            "true", "1", "yes", "sí", "si"
        ]
    else:
        use_line_length_scaling = bool(use_line_length_scaling_raw)

    print("comprobación df settigns")
    print(df_SYS_settings)


    status, condition, hydro_soc_mode = solve_opf(
        grid,
        solver_name=solver_name,
        battery_specs=battery_specs,
        final_hydro_soc_fraction=final_soc_fraction,
        initial_hydro_soc_fraction=initial_soc_fraction,
        mip_rel_gap=mip_rel_gap,
        time_limit=time_limit,
        threads=threads,
        mip_focus=mip_focus,
        method=method,
        crossover=crossover,
        numeric_focus=numeric_focus,
        bar_conv_tol=bar_conv_tol,
        feasibility_tol=feasibility_tol,
        optimality_tol=optimality_tol,
        bar_homogeneous=bar_homogeneous,
        line_flow_penalty=line_flow_penalty,
        use_line_length_scaling=use_line_length_scaling,
        additional_extra_functionality=intermediate_extra_functionality,
    )

    if status != "ok" or condition not in ["optimal", "suboptimal"]:
        print("La optimización no ha generado una solución válida. Se cancela el postprocesado.")
        return
    
    solver_results = pd.DataFrame({
        "value": {
            "objective": grid.objective,
            "status": status,
            "condition": condition,
            "hydro_soc_mode": hydro_soc_mode,
        }
    })

    report_progress(90, "optimización terminada")

    
    if horizon == "Multiperiod":
        df_solar_available = df_solar_node_profiles.set_index("time").copy()
        df_wind_available = df_wind_node_profiles.set_index("time").copy()

        df_solar_available.columns = [f"PV_{c}" for c in df_solar_available.columns]
        df_wind_available.columns = [f"Wind_{c}" for c in df_wind_available.columns]

        df_available_renewable = pd.concat(
            [df_solar_available, df_wind_available],
            axis=1
        )

        df_available_renewable.index = pd.to_datetime(df_available_renewable.index)
        df_available_renewable.index.name = "snapshot"

        export_multiperiod_results(grid, df_SYS_settings, df_available_renewable, CFsolar, CFwind, df_StorageUnit, solver_results)
    elif horizon == "Static":
        export_static_results(grid)
    else:
        raise ValueError(f"Valor de horizonte no reconocido: {horizon}")

    report_progress(100, "resultados exportados")

    return input_path
    

def main_cli():
    used_file = run_program()
    print(f"Programa ejecutado correctamente.\nArchivo usado: {used_file}")


if __name__ == "__main__":
    import traceback

    try:
        main_cli()
        input("Pulsa Enter para cerrar...")
    except Exception:
        print("Error durante la ejecución:")
        print(traceback.format_exc())
        input("Pulsa Enter para cerrar...")
