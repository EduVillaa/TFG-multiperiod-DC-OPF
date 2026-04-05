import pandas as pd
import numpy as np
import pandas as pd


def get_e_nom(grid):
    stores = grid.stores.loc[grid.stores.index.str.startswith("BatteryStore_")]

    if "e_nom_opt" in stores.columns:
        return stores["e_nom_opt"].fillna(stores["e_nom"])
    else:
        return stores["e_nom"]

def get_p_nom(grid):
    links = grid.links.loc[
        grid.links.index.str.startswith("BatteryDischarge_")
    ]

    if "p_nom_opt" in links.columns:
        return links["p_nom_opt"].fillna(links["p_nom"])
    else:
        return links["p_nom"]

def get_snapshot_hours(grid):
    """
    Devuelve una serie con la duración de cada snapshot en horas.
    Si existe snapshot_weightings, usa la columna 'objective'.
    """
    if hasattr(grid, "snapshot_weightings") and "objective" in grid.snapshot_weightings.columns:
        return grid.snapshot_weightings["objective"]
    else:
        return pd.Series(1.0, index=grid.snapshots)

def get_battery_sizes(grid):
    """
    Devuelve un DataFrame con KPIs de las baterías:
    - Energy (MWh)
    - Power (MW)
    - Duration (h)
    - Throughput (MWh)
    - Equivalent cycles
    - Real efficiency
    - Utilization factor
    - Hours SOC 0–5%
    - Hours SOC 95–100%
    """

    # --- capacidades ---
    e_nom = get_e_nom(grid)
    p_nom = get_p_nom(grid)

    # Solo stores de batería
    store_names = e_nom.index.tolist()

    # Duración de snapshots
    snapshot_hours = get_snapshot_hours(grid)

    # -----------------------------
    # SOC (%) usando stores_t.e
    # -----------------------------
    if hasattr(grid, "stores_t") and hasattr(grid.stores_t, "e") and not grid.stores_t.e.empty:
        store_energy = grid.stores_t.e.copy()

        # quedarnos solo con stores que estén en e_nom
        valid_store_names = [name for name in store_names if name in store_energy.columns]

        if valid_store_names:
            soc_percent = store_energy[valid_store_names].divide(e_nom[valid_store_names], axis=1) * 100

            LOW_THRESHOLD = 5
            HIGH_THRESHOLD = 95

            hours_empty = (soc_percent <= LOW_THRESHOLD).sum()
            hours_full = (soc_percent >= HIGH_THRESHOLD).sum()
        else:
            hours_empty = pd.Series(dtype=float)
            hours_full = pd.Series(dtype=float)
    else:
        hours_empty = pd.Series(dtype=float)
        hours_full = pd.Series(dtype=float)

    rows = []

    for store_name in store_names:
        suffix = store_name.replace("BatteryStore_", "")
        charge_link = f"BatteryCharge_{suffix}"
        discharge_link = f"BatteryDischarge_{suffix}"

        energy_mwh = float(e_nom.loc[store_name])
        power_mw = float(p_nom.loc[discharge_link]) if discharge_link in p_nom.index else np.nan

        # -----------------------------
        # Energía de carga y descarga en lado AC
        # -----------------------------
        if charge_link in grid.links_t.p0.columns:
            charge_series = grid.links_t.p0[charge_link].clip(lower=0)
            charge_mwh = float((charge_series * snapshot_hours).sum())
        else:
            charge_mwh = 0.0

        if discharge_link in grid.links_t.p1.columns:
            discharge_series = (-grid.links_t.p1[discharge_link]).clip(lower=0)
            discharge_mwh = float((discharge_series * snapshot_hours).sum())
        else:
            discharge_mwh = 0.0

        # Throughput total
        throughput_mwh = charge_mwh + discharge_mwh

        # Ciclos equivalentes
        if energy_mwh > 0:
            equivalent_cycles = throughput_mwh / (2 * energy_mwh)
        else:
            equivalent_cycles = np.nan

        # Eficiencia real
        if charge_mwh > 0:
            real_efficiency = discharge_mwh / charge_mwh
        else:
            real_efficiency = np.nan

        # Factor de utilización
        total_hours = float(snapshot_hours.sum())
        if power_mw > 0 and total_hours > 0:
            utilization_factor = discharge_mwh / (power_mw * total_hours)
        else:
            utilization_factor = np.nan

        # Horas en extremos de SOC
        h_empty = float(hours_empty.get(store_name, np.nan))
        h_full = float(hours_full.get(store_name, np.nan))

        rows.append({
            "name": store_name,
            "Energy (MWh)": energy_mwh,
            "Power (MW)": power_mw,
            "Duration (h)": energy_mwh / power_mw if power_mw > 0 else np.nan,
            "Battery charge (MWh)": round(charge_mwh, 1),
            "Battery discharge (MWh)": round(discharge_mwh, 1),
            "Throughput (MWh)": round(throughput_mwh, 1),
            "Equivalent cycles": round(equivalent_cycles, 1),
            "Real efficiency": round(real_efficiency, 4),
            "Utilization factor": round(utilization_factor, 4),
            "Hours SOC 0–5%": h_empty,
            "Hours SOC 95–100%": h_full,
        })

    df = pd.DataFrame(rows).set_index("name")

    return df