from __future__ import annotations
import pandas as pd
import pypsa
import matplotlib.pyplot as plt
import networkx as nx
from pathlib import Path
from typing import Optional, Dict, Any


def leerhojas(filename: str) -> dict:

    sheets = {}

    # --- SYS SETTINGS ---
    sheets["SYS_settings"] = pd.read_excel(
        filename,
        sheet_name="SYS_settings",
        header=1
    )

    # --- NET BUSES ---
    sheets["Net_Buses"] = pd.read_excel(
        filename,
        sheet_name="Net_Buses",
        header=1   
    ).iloc[:, 1:]  

    # --- NET LINES ---
    sheets["Net_Lines"] = pd.read_excel(
        filename,
        sheet_name="Net_Lines",
        header=2
    ).iloc[:, 1:]

    # --- NET LOADS ---
    sheets["Net_Loads"] = pd.read_excel(
        filename,
        sheet_name="Net_Loads",
        header=2
    ).iloc[:, 1:]

    # --- GEN DISPATCHABLE ---
    sheets["Gen_Dispatchable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Dispatchable",
        header=2
    ).iloc[:, 1:]

    # --- GEN RENEWABLE ---
    sheets["Gen_Renewable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Renewable",
        header=2
    ).iloc[:, 1:]

    # --- STORAGE UNIT ---
    sheets["StorageUnit"] = pd.read_excel(
        filename,
        sheet_name="StorageUnit",
        header=2
    ).iloc[:, 1:]

    return sheets

def add_storage_unit(grid: pypsa.Network, df_StorageUnit: pd.DataFrame) -> None:
    df_StorageUnit["efficiency_store (p.u)"] = pd.to_numeric(df_StorageUnit["efficiency_store (p.u)"], errors="coerce").fillna(0.95).astype(float)
    df_StorageUnit["efficiency_dispatch (p.u)"] = pd.to_numeric(df_StorageUnit["efficiency_dispatch (p.u)"], errors="coerce").fillna(0.95).astype(float)
    df_StorageUnit["standing_loss (%/h)"] = pd.to_numeric(df_StorageUnit["standing_loss (%/h)"], errors="coerce").fillna(0)
    df_StorageUnit["cyclic SOC (0/1)"] = pd.to_numeric(df_StorageUnit["cyclic SOC (0/1)"], errors="coerce").fillna(1)
    df_StorageUnit["initial SOC (%)"] = pd.to_numeric(df_StorageUnit["initial SOC (%)"], errors="coerce").fillna(0.5)
    df_StorageUnit["marginal_cost (€/MWh)"] = pd.to_numeric(df_StorageUnit["marginal_cost (€/MWh)"], errors="coerce").fillna(0)

    for n in range(df_StorageUnit["STORAGE UNIT LOCATION"].count()):
        location = df_StorageUnit.loc[n, "STORAGE UNIT LOCATION"]
        if pd.notna(location):
            p_nom = df_StorageUnit.loc[n, "Rated active power (MW)"]
            max_hours = df_StorageUnit.loc[n, "Max hours at rated active power (h)"]
            energy_capacity = p_nom * max_hours
            grid.add("StorageUnit", f"StorageUnit{location}_g{n}", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                    bus = f"Bus_node_{location}", 
                    p_nom = p_nom,  #potencia máxima de carga/descarga (MW)
                    max_hours = max_hours, #¿Cuántas horas puede descargar la batería a potencia máxima antes de vaciarse? #energía (MWh) = p_nom (MW) * max_hours (h)
                    efficiency_store = df_StorageUnit.loc[n, "efficiency_store (p.u)"], #Eficiencia del proceso de carga.
                    efficiency_dispatch = df_StorageUnit.loc[n, "efficiency_dispatch (p.u)"], #Eficiencia de la descarga.
                    standing_loss = df_StorageUnit.loc[n, "standing_loss (%/h)"], #Pérdidas por autodescarga del almacenamiento 
                    #E_{t+1} = E_t * (1 - standing_loss) 
                    # #En PyPSA se debe introducir con las unidades de p.u/snapshot
                    state_of_charge_initial = df_StorageUnit.loc[n, "initial SOC (%)"] * energy_capacity, #Energía almacenada al inicio de la simulación. PyPSA debe recibir MWh
                    cyclic_state_of_charge = df_StorageUnit.loc[n, "cyclic SOC (0/1)"], #Impone SOC_{final} = SOC_{inicial} #Esto evita que el optimizador descargue toda la batería al final o la cargue gratis.
                    marginal_cost = df_StorageUnit.loc[n, "marginal_cost (€/MWh)"], #Puede representar degradación y costes operativos
                    carrier = "AC",
                )

def build_network(df_SYS_settings: pd.DataFrame) -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")

    time_horizon = str(df_SYS_settings.loc[3, "SYSTEM PARAMETERS"])
    if time_horizon == "Static":
        grid.set_snapshots(pd.DatetimeIndex(["2026-01-01 00:00"]))

    elif time_horizon == "Day":
        grid.set_snapshots(pd.date_range("2026-01-01", periods=24, freq="h"))

    elif time_horizon == "Week":
        grid.set_snapshots(pd.date_range("2026-01-01", periods=168, freq="h"))
  
    return grid

def add_buses(grid: pypsa.Network, df_Net_Buses: pd.DataFrame) -> None:
    n_buses = df_Net_Buses["Bus rated voltage (kV)"].count()
    for n in range(n_buses):
        grid.add("Bus", f"Bus_node_{n+1}", v_nom=df_Net_Buses.loc[n, "Bus rated voltage (kV)"], carrier="AC")

def add_dispatchable_generators(grid: pypsa.Network, df_Gen_Dispatchable: pd.DataFrame) -> None:

    df_Gen_Dispatchable["Pmin (MW)"] = pd.to_numeric(df_Gen_Dispatchable["Pmin (MW)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["a (€/MW²h)"] = pd.to_numeric(df_Gen_Dispatchable["a (€/MW²h)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["b (€/MWh)"] = pd.to_numeric(df_Gen_Dispatchable["b (€/MWh)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["c (€)"] = pd.to_numeric(df_Gen_Dispatchable["c (€)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["pwl segments"] = pd.to_numeric(df_Gen_Dispatchable["pwl segments"], errors="coerce").fillna(1).astype(int)

    for n in range(df_Gen_Dispatchable["GENERATOR LOCATION"].count()):
        Pmax = float(df_Gen_Dispatchable.loc[n, "Rated active power (MW)"])
        if pd.isna(Pmax):
            continue
        location = int(df_Gen_Dispatchable.loc[n, "GENERATOR LOCATION"])
        Pmin = float(df_Gen_Dispatchable.loc[n, "Pmin (MW)"])
        segs = int(df_Gen_Dispatchable.loc[n, "pwl segments"]) if pd.notna(df_Gen_Dispatchable.loc[n, "pwl segments"]) else 1

        a = df_Gen_Dispatchable.loc[n, "a (€/MW²h)"]
        b = df_Gen_Dispatchable.loc[n, "b (€/MWh)"]

        if segs > 1:
            step = Pmax / segs
            remaining_min = Pmin

            for i in range(segs):
                block_min_mw = max(0.0, min(step, remaining_min))
                remaining_min -= block_min_mw

                p_min_pu = block_min_mw / step  # p.u. del bloque

                P_mid = (i + 0.5) * step
                marginal_cost = 2 * a * P_mid + b

                grid.add(
                    "Generator", f"DispatchGen{location}_g{n}_seg{i+1}",
                    bus=f"Bus_node_{location}",
                    p_nom=step,
                    p_min_pu=p_min_pu,
                    marginal_cost=marginal_cost,
                    carrier="AC"
                )
        else:
            grid.add(
                "Generator", f"DispatchGen{location}_g{n}_seg1", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                bus=f"Bus_node_{location}",
                p_nom=Pmax,
                p_min_pu=(Pmin / Pmax) if Pmax > 0 else 0.0,
                marginal_cost=b,
                carrier="AC"
            )

def add_loads(grid: pypsa.Network, df_Net_Loads: pd.DataFrame, df_SYS_settings: pd.DataFrame) -> None:
    df_Net_Loads["Loss factor (%)"] = pd.to_numeric(df_Net_Loads["Loss factor (%)"], errors="coerce").fillna(0)
    for n in range(df_Net_Loads["Active power demand (MW)"].last_valid_index() + 1):
        location = df_Net_Loads.loc[n, "LOAD LOCATION"]
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]
        Ploss = df_Net_Loads.loc[n, "Loss factor (%)"]
        VOLL = df_SYS_settings.loc[0, "SYSTEM PARAMETERS"] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{location}_L{n}",  #L{n} permite distinguir cargas del mismo nodo
                    bus=f"Bus_node_{location}", 
                    p_set=Pd*(1+Ploss), carrier="AC")
            
            use_shed = int(df_SYS_settings.loc[1, "SYSTEM PARAMETERS"]) == 1
            if use_shed:
                grid.add("Generator", f"shedding_gen_node_{location}", bus=f"Bus_node_{location}", 
                        p_nom=1e6, 
                        marginal_cost=VOLL,
                        p_min_pu=0,
                        carrier="AC")

def add_lines(grid: pypsa.Network, df_Net_Lines: pd.DataFrame) -> None:
    for n in range(df_Net_Lines["From"].count()):
        desde = int(df_Net_Lines.loc[n, "From"])
        hasta = int(df_Net_Lines.loc[n, "To"])
        grid.add(
            "Line", f"L{desde}{hasta}",
            bus0=f"Bus_node_{desde}",
            bus1=f"Bus_node_{hasta}",
            x=df_Net_Lines.loc[n, "Reactance (p.u)"],
            r=1e-6, #Para evitar el warning que sale al no incluir la resistencia
            s_nom=df_Net_Lines.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )

def add_renewable_generator(grid: pypsa.Network, df_Gen_Renewable: pd.DataFrame, df_SYS_settings: pd.DataFrame) -> None:

    time_horizon = str(df_SYS_settings.loc[3, "SYSTEM PARAMETERS"])

    if time_horizon == "Static":
            pv_profile = pd.Series(
            [1],
            index=grid.snapshots)

    elif time_horizon == "Day":
        pv_profile = [0, 0, 0, 0, 0, 0,
        0.05, 0.15, 0.35, 0.60, 0.80, 0.95,
        1.00, 0.90, 0.70, 0.45, 0.20, 0.05,
        0, 0, 0, 0, 0, 0]

    elif time_horizon == "Week":
        pv_profile = (
        [0,0,0,0,0,0,0.05,0.15,0.35,0.60,0.80,0.95,1.0,0.9,0.7,0.45,0.2,0.05,0,0,0,0,0,0] +   # Mon sunny
        [0,0,0,0,0,0,0.04,0.12,0.30,0.55,0.70,0.80,0.85,0.75,0.55,0.35,0.15,0.04,0,0,0,0,0,0] + # Tue cloudy
        [0,0,0,0,0,0,0.05,0.18,0.40,0.70,0.90,1.0,1.0,0.95,0.80,0.55,0.25,0.08,0,0,0,0,0,0] +   # Wed clear
        [0,0,0,0,0,0,0.03,0.10,0.25,0.45,0.60,0.70,0.75,0.65,0.45,0.25,0.10,0.03,0,0,0,0,0,0] + # Thu cloudy
        [0,0,0,0,0,0,0.05,0.20,0.45,0.75,0.95,1.0,1.0,0.9,0.75,0.50,0.25,0.08,0,0,0,0,0,0] +    # Fri clear
        [0,0,0,0,0,0,0.04,0.15,0.30,0.55,0.75,0.85,0.90,0.80,0.60,0.40,0.18,0.05,0,0,0,0,0,0] + # Sat mixed
        [0,0,0,0,0,0,0.02,0.08,0.20,0.35,0.50,0.60,0.65,0.55,0.40,0.25,0.10,0.02,0,0,0,0,0,0]   # Sun bad weather
        )

    for n in range(df_Gen_Renewable["GENERATOR LOCATION"].count()):
        location = df_Gen_Renewable.loc[n, "GENERATOR LOCATION"]
        if pd.notna(location):
            grid.add("Generator", f"RenewableGen{location}_g{n}", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                    bus = f"Bus_node_{location}", 
                    p_nom = df_Gen_Renewable.loc[n, "Rated active power (MW)"],
                    p_min_pu = 0, #No hay restricción de potencia mínima para las renovables
                    marginal_cost = 0, #El coste marginal para las renovables se considera nulo
                    carrier= "AC")

            grid.generators_t.p_max_pu[f"RenewableGen{location}_g{n}"] = pv_profile

def solve_opf(grid: pypsa.Network, solver_name) -> None:
    grid.optimize(solver_name=solver_name)

def export_results(
    n: pypsa.Network,
    out_path: str | Path,
    *,
    case_name: Optional[str] = None,
    include_prices: bool = True,
    include_debug: bool = False,
) -> Path:
    """
    Exporta resultados de PyPSA a un Excel de informe, con varias hojas ordenadas.

    Hojas que genera (si existen datos):
      - 00_Summary      (KPIs y metadatos)
      - 10_Balance      (carga, generación, storage neto, ENS, etc.)
      - 20_Dispatch     (potencias de generadores por snapshot)
      - 30_Renewables   (gen renovable + curtailment si hay p_max_pu)
      - 40_LineFlows    (flujos por línea y % loading si hay s_nom)
      - 50_Storage      (p, SOC si existe storage_units)
      - 60_Prices       (marginal prices por bus, si existe y include_prices)
      - 90_Debug        (opcional: tablas auxiliares)

    Nota: crea un archivo NUEVO (no modifica tu Excel de inputs).
    """
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    # ---------- helpers ----------
    def _safe_df(x) -> Optional[pd.DataFrame]:
        if x is None:
            return None
        if isinstance(x, pd.Series):
            return x.to_frame()
        if isinstance(x, (pd.DataFrame,)):
            return x
        # escalar
        return pd.DataFrame({"value": [x]})

    def _has(attr: str) -> bool:
        try:
            obj = getattr(n, attr)
            return obj is not None
        except Exception:
            return False

    snapshots = pd.Index(n.snapshots, name="snapshot")

    # ---------- KPIs ----------
    kpi_rows = []

    # Objective / cost (si has corrido lopf/opf)
    obj = getattr(n, "objective", None)
    if obj is not None:
        kpi_rows.append(("objective", float(obj), "€ (si tus costes están en €)"))

    # Total load energy (MWh) si snapshots son horarias (o ponderadas)
    # Usamos snapshot_weightings si existe; si no, peso=1
    if hasattr(n, "snapshot_weightings") and "objective" in n.snapshot_weightings.columns:
        w = n.snapshot_weightings["objective"].reindex(snapshots).fillna(1.0)
    elif hasattr(n, "snapshot_weightings") and "generators" in n.snapshot_weightings.columns:
        w = n.snapshot_weightings["generators"].reindex(snapshots).fillna(1.0)
    else:
        w = pd.Series(1.0, index=snapshots, name="weight")

    # Load (p_set o p) por snapshot
    load_ts = None
    if hasattr(n, "loads_t") and hasattr(n.loads_t, "p_set") and len(getattr(n, "loads", [])) > 0:
        load_ts = n.loads_t.p_set.reindex(snapshots)
    elif hasattr(n, "loads_t") and hasattr(n.loads_t, "p") and len(getattr(n, "loads", [])) > 0:
        load_ts = n.loads_t.p.reindex(snapshots)

    if load_ts is not None and not load_ts.empty:
        total_load_mwh = (load_ts.sum(axis=1) * w).sum()
        peak_load_mw = load_ts.sum(axis=1).max()
        kpi_rows += [
            ("total_load_energy", float(total_load_mwh), "MWh (según ponderación)"),
            ("peak_load", float(peak_load_mw), "MW"),
        ]

    # Generator dispatch
    gen_p = None
    if hasattr(n, "generators_t") and hasattr(n.generators_t, "p") and len(getattr(n, "generators", [])) > 0:
        gen_p = n.generators_t.p.reindex(snapshots)

    if gen_p is not None and not gen_p.empty:
        total_gen_mwh = (gen_p.sum(axis=1) * w).sum()
        kpi_rows.append(("total_generation", float(total_gen_mwh), "MWh (según ponderación)"))

    # Storage
    su_p = None
    su_soc = None
    if hasattr(n, "storage_units_t") and len(getattr(n, "storage_units", [])) > 0:
        if hasattr(n.storage_units_t, "p"):
            su_p = n.storage_units_t.p.reindex(snapshots)
        if hasattr(n.storage_units_t, "state_of_charge"):
            su_soc = n.storage_units_t.state_of_charge.reindex(snapshots)

    # Load shedding (si modelas un generador/elemento de shed; aquí detectamos por nombre típico)
    # Si tienes una convención distinta, puedes ajustar este bloque.
    ens_mwh = None
    if gen_p is not None and not gen_p.empty:
        shed_cols = [c for c in gen_p.columns if str(c).lower().find("shed") >= 0]
        if shed_cols:
            ens_mwh = ((gen_p[shed_cols].sum(axis=1) * w).sum())
            kpi_rows.append(("energy_not_served_ENS", float(ens_mwh), "MWh"))

    # Curtailment renovable (si existe p_max_pu)
    curtail_mwh = None
    if gen_p is not None and hasattr(n.generators_t, "p_max_pu") and hasattr(n, "generators"):
        p_max_pu = n.generators_t.p_max_pu.reindex(snapshots)
        # capacidad nominal
        p_nom = n.generators.p_nom.reindex(gen_p.columns).fillna(0.0)
        available = p_max_pu.mul(p_nom, axis=1)
        common = [c for c in gen_p.columns if c in available.columns]
        if common:
            curtail = (available[common] - gen_p[common]).clip(lower=0.0)
            curtail_mwh = (curtail.sum(axis=1) * w).sum()
            kpi_rows.append(("renewable_curtailment", float(curtail_mwh), "MWh"))

    summary_df = pd.DataFrame(kpi_rows, columns=["KPI", "Value", "Unit"])
    if case_name:
        meta_top = pd.DataFrame(
            {
                "Field": ["case_name", "n_snapshots"],
                "Value": [case_name, int(len(snapshots))],
            }
        )
    else:
        meta_top = pd.DataFrame(
            {
                "Field": ["n_snapshots"],
                "Value": [int(len(snapshots))],
            }
        )

    # ---------- Balance ----------
    # Construimos una tabla por snapshot: Load, Gen, Storage_net (descarga positiva), ENS, etc.
    balance = pd.DataFrame(index=snapshots)

    if load_ts is not None and not load_ts.empty:
        balance["load_MW"] = load_ts.sum(axis=1)

    if gen_p is not None and not gen_p.empty:
        balance["generation_MW"] = gen_p.sum(axis=1)

    if su_p is not None and not su_p.empty:
        # en PyPSA: storage_units_t.p >0 suele ser descarga a la red, <0 carga (según convención)
        balance["storage_net_MW"] = su_p.sum(axis=1)

    if ens_mwh is not None:
        # ENS por snapshot (MW) si existe shedding
        shed_cols = [c for c in gen_p.columns if str(c).lower().find("shed") >= 0]
        balance["load_shedding_MW"] = gen_p[shed_cols].sum(axis=1)

    if curtail_mwh is not None:
        # curtailment por snapshot (MW)
        curtail_cols = None
        try:
            curtail_cols = curtail.columns  # type: ignore[name-defined]
        except Exception:
            curtail_cols = None
        if curtail_cols is not None:
            balance["curtailment_MW"] = curtail.sum(axis=1)  # type: ignore[name-defined]

    # ---------- Dispatch ----------
    dispatch_df = None
    if gen_p is not None and not gen_p.empty:
        dispatch_df = gen_p.copy()
        dispatch_df.index.name = "snapshot"

    # ---------- Renewables ----------
    renew_df = None
    if gen_p is not None and not gen_p.empty:
        # Heurística: renovables = control no importa; mejor por carrier si está definido
        if hasattr(n, "generators") and "carrier" in n.generators.columns:
            ren_mask = n.generators["carrier"].astype(str).str.lower().isin(
                ["pv", "solar", "wind", "onwind", "offwind", "ror", "hydro", "renewable"]
            )
            ren_gens = n.generators.index[ren_mask].intersection(gen_p.columns)
            if len(ren_gens) > 0:
                renew_df = gen_p[ren_gens].copy()
        # Si no hay carrier, no forzamos.

    # Curtailment detallado si existe
    curtail_detail = None
    if gen_p is not None and hasattr(n.generators_t, "p_max_pu") and hasattr(n, "generators"):
        p_max_pu = n.generators_t.p_max_pu.reindex(snapshots)
        p_nom = n.generators.p_nom.reindex(gen_p.columns).fillna(0.0)
        available = p_max_pu.mul(p_nom, axis=1)
        common = [c for c in gen_p.columns if c in available.columns]
        if common:
            curtail_detail = (available[common] - gen_p[common]).clip(lower=0.0)
            curtail_detail.index.name = "snapshot"

    # ---------- Line flows ----------
    lineflows_df = None
    lineload_df = None
    if hasattr(n, "lines_t") and hasattr(n.lines_t, "p0") and len(getattr(n, "lines", [])) > 0:
        lineflows_df = n.lines_t.p0.reindex(snapshots).copy()
        lineflows_df.index.name = "snapshot"

        # % loading si hay s_nom
        if hasattr(n, "lines") and "s_nom" in n.lines.columns:
            s_nom = n.lines.s_nom.reindex(lineflows_df.columns).replace(0, pd.NA)
            lineload_df = (lineflows_df.abs().div(s_nom, axis=1) * 100.0)
            lineload_df.index.name = "snapshot"

    # ---------- Storage ----------
    storage_p_df = None
    storage_soc_df = None
    if su_p is not None and not su_p.empty:
        storage_p_df = su_p.copy()
        storage_p_df.index.name = "snapshot"
    if su_soc is not None and not su_soc.empty:
        storage_soc_df = su_soc.copy()
        storage_soc_df.index.name = "snapshot"

    # ---------- Prices ----------
    prices_df = None
    if include_prices and hasattr(n, "buses_t") and hasattr(n.buses_t, "marginal_price"):
        mp = n.buses_t.marginal_price.reindex(snapshots)
        if mp is not None and not mp.empty:
            prices_df = mp.copy()
            prices_df.index.name = "snapshot"

    # ---------- Write Excel ----------
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
        # 00 Summary: metadatos + KPIs
        meta_top.to_excel(writer, sheet_name="00_Summary", index=False, startrow=0)
        summary_df.to_excel(writer, sheet_name="00_Summary", index=False, startrow=len(meta_top) + 2)

        # 10 Balance
        if not balance.empty:
            balance.to_excel(writer, sheet_name="10_Balance")

        # 20 Dispatch
        if dispatch_df is not None and not dispatch_df.empty:
            dispatch_df.to_excel(writer, sheet_name="20_Dispatch")

        # 30 Renewables
        if renew_df is not None and not renew_df.empty:
            renew_df.to_excel(writer, sheet_name="30_Renewables", startrow=0)
            if curtail_detail is not None and not curtail_detail.empty:
                # pegamos curtailment debajo, con un título
                start = len(renew_df) + 3
                pd.DataFrame({"NOTE": ["Curtailment (available - dispatched), MW"]}).to_excel(
                    writer, sheet_name="30_Renewables", index=False, startrow=start
                )
                curtail_detail.to_excel(writer, sheet_name="30_Renewables", startrow=start + 2)
        elif curtail_detail is not None and not curtail_detail.empty:
            curtail_detail.to_excel(writer, sheet_name="30_Renewables")

        # 40 Line flows
        if lineflows_df is not None and not lineflows_df.empty:
            lineflows_df.to_excel(writer, sheet_name="40_LineFlows", startrow=0)
            if lineload_df is not None and not lineload_df.empty:
                start = len(lineflows_df) + 3
                pd.DataFrame({"NOTE": ["Line loading (%), based on |p0| / s_nom * 100"]}).to_excel(
                    writer, sheet_name="40_LineFlows", index=False, startrow=start
                )
                lineload_df.to_excel(writer, sheet_name="40_LineFlows", startrow=start + 2)

        # 50 Storage
        if (storage_p_df is not None and not storage_p_df.empty) or (storage_soc_df is not None and not storage_soc_df.empty):
            sheet = "50_Storage"
            r = 0
            if storage_p_df is not None and not storage_p_df.empty:
                pd.DataFrame({"NOTE": ["Storage power p (MW). Sign convention: >0 discharge to grid, <0 charge (typical in PyPSA)."]}).to_excel(
                    writer, sheet_name=sheet, index=False, startrow=r
                )
                storage_p_df.to_excel(writer, sheet_name=sheet, startrow=r + 2)
                r = r + 2 + len(storage_p_df) + 3
            if storage_soc_df is not None and not storage_soc_df.empty:
                pd.DataFrame({"NOTE": ["State of charge (MWh)"]}).to_excel(
                    writer, sheet_name=sheet, index=False, startrow=r
                )
                storage_soc_df.to_excel(writer, sheet_name=sheet, startrow=r + 2)

        # 60 Prices
        if prices_df is not None and not prices_df.empty:
            prices_df.to_excel(writer, sheet_name="60_Prices")

        # 90 Debug (opcional)
        if include_debug:
            # ejemplo: exportar tablas estáticas relevantes
            if hasattr(n, "generators") and len(n.generators) > 0:
                n.generators.to_excel(writer, sheet_name="90_Debug", startrow=0)
                r = len(n.generators) + 3
            else:
                r = 0
            if hasattr(n, "lines") and len(getattr(n, "lines", [])) > 0:
                n.lines.to_excel(writer, sheet_name="90_Debug", startrow=r)

    return out_path


# ------------------ ejemplo de uso ------------------
# n.lopf(...) o n.optimize(...) primero
# export_results(n, "ExampleGrid_RESULTS.xlsx", case_name="ExampleGrid - Week", include_prices=True)

def drawGrid(grid: pypsa.Network):
    # convertir red PyPSA a grafo
    G = grid.graph()

    # generar layout automático
    #pos = nx.spring_layout(G)
    pos = nx.circular_layout(G)
    #pos = nx.kamada_kawai_layout(G)

    # asignar coordenadas a buses automáticamente
    for bus in grid.buses.index:
        grid.buses.loc[bus, "x"] = pos[bus][0]
        grid.buses.loc[bus, "y"] = pos[bus][1]

    # dibujar
    grid.plot()
    plt.show()

def main():
    data = leerhojas("ExampleGrid.xlsx")

    df_SYS_settings = data["SYS_settings"]
    df_Net_Buses = data["Net_Buses"]
    df_Net_Lines = data["Net_Lines"]
    df_Net_Loads = data["Net_Loads"]
    df_Gen_Dispatchable = data["Gen_Dispatchable"]
    df_Gen_Renewable = data["Gen_Renewable"]
    df_StorageUnit = data["StorageUnit"]
    
    grid = build_network(df_SYS_settings)

    add_buses(grid, df_Net_Buses)
    add_lines(grid, df_Net_Lines)
    add_loads(grid, df_Net_Loads, df_SYS_settings) #Puesto que incluimos los generadores VOLL allí donde haya cargas 
    #la función add_loads debe recibir también df_SYS_settings para leer el VOLL en €/MWh introducido por el usuario

    add_dispatchable_generators(grid, df_Gen_Dispatchable)
    add_renewable_generator(grid, df_Gen_Renewable, df_SYS_settings)
    add_storage_unit(grid, df_StorageUnit)

    solver=str(df_SYS_settings.loc[2, "SYSTEM PARAMETERS"])
    solve_opf(grid, solver_name=solver)
    #print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])
    #print(grid.storage_units[["p_nom", "max_hours", "bus", "efficiency_store", "efficiency_dispatch", "standing_loss", "state_of_charge_initial", "cyclic_state_of_charge", "marginal_cost"]])
    #print("\n")

    export_results(grid, "ExampleGrid_RESULTS.xlsx", case_name="ExampleGrid - Week", include_prices=True)




if __name__ == "__main__":
    main()




