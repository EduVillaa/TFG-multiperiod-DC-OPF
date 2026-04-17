import pandas as pd
import numpy as np


def get_e_nom(grid: object) -> pd.Series:
    """
    Devuelve la capacidad energética de las baterías (MWh).

    Prioriza e_nom_opt si existe; si no, usa e_nom.
    Índice: nombres de los BatteryStore_*.
    """
    if not hasattr(grid, "stores") or grid.stores is None or grid.stores.empty:
        return pd.Series(dtype=float)

    stores = grid.stores.loc[
        grid.stores.index.astype(str).str.startswith("BatteryStore_")
    ].copy()

    if stores.empty:
        return pd.Series(dtype=float)

    if "e_nom_opt" in stores.columns:
        return stores["e_nom_opt"].fillna(stores.get("e_nom"))
    return stores["e_nom"]


def get_p_nom(grid: object) -> pd.Series:
    """
    Devuelve la potencia nominal de las baterías (MW).

    Se toma del link de descarga BatteryDischarge_*.
    Prioriza p_nom_opt si existe; si no, usa p_nom.
    """
    if not hasattr(grid, "links") or grid.links is None or grid.links.empty:
        return pd.Series(dtype=float)

    links = grid.links.loc[
        grid.links.index.astype(str).str.startswith("BatteryDischarge_")
    ].copy()

    if links.empty:
        return pd.Series(dtype=float)

    if "p_nom_opt" in links.columns:
        return links["p_nom_opt"].fillna(links.get("p_nom"))
    return links["p_nom"]


def get_snapshot_hours(grid: object) -> pd.Series:
    """
    Devuelve una serie con la duración de cada snapshot en horas.
    Si existe snapshot_weightings, usa la columna 'objective'.
    """
    if (
        hasattr(grid, "snapshot_weightings")
        and grid.snapshot_weightings is not None
        and "objective" in grid.snapshot_weightings.columns
    ):
        return grid.snapshot_weightings["objective"].reindex(grid.snapshots).fillna(1.0)

    return pd.Series(1.0, index=grid.snapshots, dtype=float)


def get_battery_sizes(grid: object) -> pd.DataFrame:
    """
    Devuelve un DataFrame con KPIs de las baterías modeladas como Store + Link:

    - Energy (MWh)
    - Power (MW)
    - Duration (h)
    - Battery charge (MWh)
    - Battery discharge (MWh)
    - Throughput (MWh)
    - Equivalent cycles
    - Real efficiency
    - Utilization factor
    - Hours SOC 0–5%
    - Hours SOC 95–100%

    La duración ya no es un input del modelo: se calcula como KPI
    a partir de Energy / Power.

    Si no hay baterías en la red, devuelve un DataFrame vacío con
    las columnas esperadas.
    """

    output_columns = [
        "Energy (MWh)",
        "Power (MW)",
        "Duration (h)",
        "Battery charge (MWh)",
        "Battery discharge (MWh)",
        "Throughput (MWh)",
        "Equivalent cycles",
        "Real efficiency",
        "Utilization factor",
        "Hours SOC 0–5%",
        "Hours SOC 95–100%",
    ]

    empty_result = pd.DataFrame(columns=output_columns)
    empty_result.index.name = "name"

    # -----------------------------
    # Comprobaciones básicas
    # -----------------------------
    if not hasattr(grid, "stores") or grid.stores is None or grid.stores.empty:
        return empty_result

    if not hasattr(grid, "links") or grid.links is None or grid.links.empty:
        return empty_result

    store_names = [
        str(name)
        for name in grid.stores.index
        if str(name).startswith("BatteryStore_")
    ]

    if not store_names:
        return empty_result

    # -----------------------------
    # Capacidades nominales
    # -----------------------------
    e_nom = get_e_nom(grid).reindex(store_names)

    discharge_links = [f"BatteryDischarge_{name.replace('BatteryStore_', '')}" for name in store_names]
    charge_links = [f"BatteryCharge_{name.replace('BatteryStore_', '')}" for name in store_names]

    discharge_p_nom = get_p_nom(grid).reindex(discharge_links)

    # Potencia de carga por si hiciera falta como fallback
    if hasattr(grid, "links") and grid.links is not None and not grid.links.empty:
        charge_links_df = grid.links.loc[
            grid.links.index.astype(str).isin(charge_links)
        ].copy()

        if "p_nom_opt" in charge_links_df.columns:
            charge_p_nom = charge_links_df["p_nom_opt"].fillna(charge_links_df.get("p_nom"))
        else:
            charge_p_nom = charge_links_df.get("p_nom", pd.Series(dtype=float))
    else:
        charge_p_nom = pd.Series(dtype=float)

    snapshot_hours = get_snapshot_hours(grid)

    # -----------------------------
    # SOC y horas cerca de vacío/lleno
    # -----------------------------
    hours_empty = pd.Series(dtype=float)
    hours_full = pd.Series(dtype=float)

    if (
        hasattr(grid, "stores_t")
        and hasattr(grid.stores_t, "e")
        and grid.stores_t.e is not None
        and not grid.stores_t.e.empty
    ):
        store_energy = grid.stores_t.e.copy()
        valid_store_names = [s for s in store_names if s in store_energy.columns]

        if valid_store_names:
            denom = e_nom.reindex(valid_store_names).replace(0.0, np.nan)
            soc_percent = store_energy[valid_store_names].divide(denom, axis=1) * 100.0

            low_threshold = 5.0
            high_threshold = 95.0

            hours_empty = (soc_percent <= low_threshold).mul(snapshot_hours, axis=0).sum(axis=0)
            hours_full = (soc_percent >= high_threshold).mul(snapshot_hours, axis=0).sum(axis=0)

    # -----------------------------
    # Series temporales de links
    # -----------------------------
    links_t_p0 = pd.DataFrame(index=grid.snapshots)
    links_t_p1 = pd.DataFrame(index=grid.snapshots)

    if hasattr(grid, "links_t") and grid.links_t is not None:
        if hasattr(grid.links_t, "p0") and grid.links_t.p0 is not None and not grid.links_t.p0.empty:
            links_t_p0 = grid.links_t.p0.copy()
        if hasattr(grid.links_t, "p1") and grid.links_t.p1 is not None and not grid.links_t.p1.empty:
            links_t_p1 = grid.links_t.p1.copy()

    # -----------------------------
    # KPIs por batería
    # -----------------------------
    rows = []

    for store_name in store_names:
        suffix = store_name.replace("BatteryStore_", "")
        charge_link = f"BatteryCharge_{suffix}"
        discharge_link = f"BatteryDischarge_{suffix}"

        energy_mwh = e_nom.get(store_name, np.nan)
        energy_mwh = float(energy_mwh) if pd.notna(energy_mwh) else np.nan

        power_mw = discharge_p_nom.get(discharge_link, np.nan)
        if pd.isna(power_mw):
            power_mw = charge_p_nom.get(charge_link, np.nan)
        power_mw = float(power_mw) if pd.notna(power_mw) else np.nan

        # Carga en lado AC: bus0 -> bus1 del link de carga
        if charge_link in links_t_p0.columns:
            charge_series = links_t_p0[charge_link].clip(lower=0.0)
            charge_mwh = float((charge_series * snapshot_hours).sum())
        else:
            charge_mwh = 0.0

        # Descarga en lado AC: energía entregada al bus AC
        # En tu convención previa estabas usando -p1
        if discharge_link in links_t_p1.columns:
            discharge_series = (-links_t_p1[discharge_link]).clip(lower=0.0)
            discharge_mwh = float((discharge_series * snapshot_hours).sum())
        else:
            discharge_mwh = 0.0

        throughput_mwh = charge_mwh + discharge_mwh

        if pd.notna(energy_mwh) and energy_mwh > 0.0:
            equivalent_cycles = throughput_mwh / (2.0 * energy_mwh)
        else:
            equivalent_cycles = np.nan

        if charge_mwh > 0.0:
            real_efficiency = discharge_mwh / charge_mwh
        else:
            real_efficiency = np.nan

        total_hours = float(snapshot_hours.sum())

        if pd.notna(energy_mwh) and pd.notna(power_mw) and power_mw > 0.0:
            duration_h = energy_mwh / power_mw
        else:
            duration_h = np.nan

        if pd.notna(power_mw) and power_mw > 0.0 and total_hours > 0.0:
            utilization_factor = discharge_mwh / (power_mw * total_hours)
        else:
            utilization_factor = np.nan

        h_empty = float(hours_empty.get(store_name, np.nan))
        h_full = float(hours_full.get(store_name, np.nan))

        rows.append(
            {
                "name": store_name,
                "Energy (MWh)": round(energy_mwh, 3) if pd.notna(energy_mwh) else np.nan,
                "Power (MW)": round(power_mw, 3) if pd.notna(power_mw) else np.nan,
                "Duration (h)": round(duration_h, 3) if pd.notna(duration_h) else np.nan,
                "Battery charge (MWh)": round(charge_mwh, 1),
                "Battery discharge (MWh)": round(discharge_mwh, 1),
                "Throughput (MWh)": round(throughput_mwh, 1),
                "Equivalent cycles": round(equivalent_cycles, 3) if pd.notna(equivalent_cycles) else np.nan,
                "Real efficiency": round(real_efficiency, 4) if pd.notna(real_efficiency) else np.nan,
                "Utilization factor": round(utilization_factor, 4) if pd.notna(utilization_factor) else np.nan,
                "Hours SOC 0–5%": round(h_empty, 3) if pd.notna(h_empty) else np.nan,
                "Hours SOC 95–100%": round(h_full, 3) if pd.notna(h_full) else np.nan,
            }
        )

    if not rows:
        return empty_result

    return pd.DataFrame(rows).set_index("name")