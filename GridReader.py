from __future__ import annotations
import pandas as pd
import pypsa
import matplotlib.pyplot as plt
import networkx as nx
from pathlib import Path
from typing import Optional, Dict, Any
import matplotlib.dates as mdates
import plotly.graph_objects as go


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
        header=3
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

    # --- GRID CONNECTION ---
    sheets["Grid_connection"] = pd.read_excel(
        filename,
        sheet_name="Grid_connection",
        header=2
    ).iloc[:, 1:]

    # --- TS WIND PROFILES ---
    sheets["TS_Wind_Profiles"] = pd.read_excel(
        filename,
        sheet_name="TS_Wind_Profiles",
        header=0
    ).iloc[:, 0:]

    # --- TS PV PROFILES ---
    sheets["TS_PV_Profiles"] = pd.read_excel(
        filename,
        sheet_name="TS_PV_Profiles",
        header=0
    ).iloc[:, 0:]

    # --- TS ENERGY PRICES PROFILES ---
    sheets["TS_Energy_Prices"] = pd.read_excel(
        filename,
        sheet_name="TS_Energy_Prices",
        header=0
    ).iloc[:, 0:]

    # --- TS LOAD PROFILES ---
    sheets["TS_LoadProfiles"] = pd.read_excel(
        filename,
        sheet_name="TS_LoadProfiles",
        header=0
    ).iloc[:, 0:]

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
            
def grid_connection(grid: pypsa.Network, df_Grid_connection: pd.DataFrame, df_TS_Energy_Prices: pd.DataFrame, df_SYS_settings: pd.DataFrame) -> None:
     
    grid.add(
        "Bus",
        "PCC",
        v_nom=df_Grid_connection.loc[0, "Grid rated voltage at the PCC"],
        carrier="AC"
    )

    df_Grid_connection["Thermal limit (MW)"] = pd.to_numeric(
        df_Grid_connection["Thermal limit (MW)"], errors="coerce"
    ).fillna(1e6) # Si se deja vacía la columna de límite térmico se asume que no hay límite.

    for n in range(df_Grid_connection["Bus"].count()):
        grid.add(
            "Line",
            f"LPCC{int(df_Grid_connection.loc[n, 'Bus'])}",
            bus0="PCC",
            bus1=f"Bus_node_{int(df_Grid_connection.loc[n, 'Bus'])}",
            x=df_Grid_connection.loc[n, "Reactance (p.u)"],
            r=1e-6,
            s_nom=df_Grid_connection.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )

    # Compra a red
    grid.add(
        "Generator",
        "Grid_import",
        bus="PCC",
        p_nom = df_Grid_connection.loc[0, "Import"],
        marginal_cost = df_Grid_connection.loc[2, "Import"], #Para OPF estático
        carrier="AC"
    )
    # Venta a red
    grid.add(
        "Generator",
        "Grid_export",
        bus="PCC",
        p_nom=df_Grid_connection.loc[0, "Export"],
        marginal_cost = -df_Grid_connection.loc[2, "Export"]*df_Grid_connection.loc[2, "Import"], #Para OPF estático
        #El coste marginal de exportación se da como porcentaje del coste marginal de importación
        sign=-1,
        carrier="AC"
    )

    #Los perfiles de tiempo temporales reescriben los estáticos si se realiza un OPF multiperiodo
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    if horizon != "Static":
        price_profile = read_prices(df_SYS_settings, df_TS_Energy_Prices)
        grid.generators_t.marginal_cost["Grid_import"] = price_profile
        grid.generators_t.marginal_cost["Grid_export"] = -price_profile*df_Grid_connection.loc[2, "Export"] #El coste marginal de exportación 
        #se da como porcentaje del coste marginal de importación

def build_network(df_SYS_settings: pd.DataFrame) -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")
    start_date = str(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])
    horizon = str(df_SYS_settings.loc[3, "SYSTEM PARAMETERS"])
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24
    if horizon == "Static":
        grid.set_snapshots(pd.DatetimeIndex(["2026-01-01 00:00"]))

    elif horizon == "Multiperiod":
        grid.set_snapshots(pd.date_range(start_date, periods=simulation_hours, freq="h"))
  
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
                    carrier="AC",
                )
        else:
            grid.add(
                "Generator", f"DispatchGen{location}_g{n}_seg1", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                bus=f"Bus_node_{location}",
                p_nom=Pmax,
                p_min_pu=(Pmin / Pmax) if Pmax > 0 else 0.0,
                marginal_cost=b,
                carrier="AC",
            )

def load_profile_reader(df_TS_LoadProfiles: pd.DataFrame, df_SYS_settings: pd.DataFrame, load_profile_type)-> pd.Series:
    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        df = df_TS_LoadProfiles.loc[start_Date:].iloc[:simulation_hours]
        return df.loc[:, load_profile_type]

def add_loads(grid: pypsa.Network, df_Net_Loads: pd.DataFrame, df_SYS_settings: pd.DataFrame, df_TS_LoadProfiles: pd.DataFrame) -> None:
    df_Net_Loads["Loss factor (%)"] = pd.to_numeric(df_Net_Loads["Loss factor (%)"], errors="coerce").fillna(0)

    for n in range(df_Net_Loads["Active power demand (MW)"].last_valid_index() + 1):
        
        if df_Net_Loads.loc[n, "Time series load profile"] == "Residencial (P2.0TD)":
            load_profile_type = "COEF. PERFIL P2.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Commercial / Industrial (P3.0TD)":
            load_profile_type = "COEF. PERFIL P3.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Electric vehicle (P3.0TDVE)":
            load_profile_type = "COEF. PERFIL P3.0TDVE"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Flat load profile":
            load_profile_type = "Flat LP"

        AnnualConsumption = df_Net_Loads.loc[n, "Annual energy consumption (MWh/year)"]
        load_profile_MW = load_profile_reader(df_TS_LoadProfiles, df_SYS_settings, load_profile_type) * AnnualConsumption
        horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
        location = df_Net_Loads.loc[n, "LOAD LOCATION"]
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]
        Ploss = df_Net_Loads.loc[n, "Loss factor (%)"]
        VOLL = df_SYS_settings.loc[0, "SYSTEM PARAMETERS"] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{location}_L{n}",  #L{n} permite distinguir cargas del mismo nodo
                    bus=f"Bus_node_{location}", 
                    p_set=Pd*(1+Ploss), carrier="AC")
            
            if horizon == "Day" or horizon == "Week":
                grid.loads_t.p_set.loc[:, f"Load_node_{location}_L{n}"] = load_profile_MW
            
            use_shed = int(df_SYS_settings.loc[1, "SYSTEM PARAMETERS"]) == 1
            if use_shed:
                grid.add("Generator", f"shedding_gen_node_{location}", bus=f"Bus_node_{location}", 
                        p_nom=1e6, 
                        marginal_cost=VOLL,
                        p_min_pu=0,
                        carrier="AC")

def add_lines(grid: pypsa.Network, df_Net_Lines: pd.DataFrame) -> None:
    df_Net_Lines["Thermal limit (MW)"] = pd.to_numeric(
        df_Net_Lines["Thermal limit (MW)"], # Si se deja vacía la columna de límite térmico se asume que no hay límite.
        errors="coerce").fillna(1e6) 
    
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

def read_prices(df_SYS_settings: pd.DataFrame,
                df_TS_Energy_Prices: pd.DataFrame) -> pd.Series:

    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        df = df_TS_Energy_Prices.loc[start_Date:].iloc[:simulation_hours]
        return df["Precio mercado SPOT Diario España (€/MWh)"]

def wind_series_reader(df_SYS_settings: pd.DataFrame,
                            df_TS_Wind_Profiles: pd.DataFrame) -> pd.Series:
    
    region = str(df_SYS_settings.loc[4, "SYSTEM PARAMETERS"])
    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24
    
    if horizon == "Static":
        return pd.Series([1], name=region)

    elif horizon == "Multiperiod":
        return df_TS_Wind_Profiles.loc[start_Date:, region].iloc[:simulation_hours]
    
def pv_series_reader(df_SYS_settings: pd.DataFrame,
                            df_TS_PV_Profiles: pd.DataFrame) -> pd.Series:
    region = str(df_SYS_settings.loc[4, "SYSTEM PARAMETERS"])
    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24

    if horizon == "Static":
        return pd.Series([1], name=region)

    elif horizon == "Multiperiod":
        return df_TS_PV_Profiles.loc[start_Date:, region].iloc[:simulation_hours]

def add_renewable_generator(grid: pypsa.Network, df_Gen_Renewable: pd.DataFrame,
    df_SYS_settings: pd.DataFrame, df_TS_Wind_Profiles: pd.DataFrame, df_TS_PV_Profiles: pd.DataFrame) -> None:

    wind_profile = wind_series_reader(df_SYS_settings, df_TS_Wind_Profiles)
    pv_profile = pv_series_reader(df_SYS_settings, df_TS_PV_Profiles)

    for n in range(df_Gen_Renewable["GENERATOR LOCATION"].count()):
        location = df_Gen_Renewable.loc[n, "GENERATOR LOCATION"]
        if pd.notna(location):
            
            if df_Gen_Renewable.loc[n, "Renewable Type"] == "PV":
                grid.add("Generator", f"PV{location}_g{n}", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                    bus = f"Bus_node_{location}", 
                    p_nom = df_Gen_Renewable.loc[n, "Rated active power (MW)"],
                    p_min_pu = 0, #No hay restricción de potencia mínima para las renovables
                    marginal_cost = 0, #El coste marginal para las renovables se considera nulo
                    carrier= "AC")
                grid.generators_t.p_max_pu[f"PV{location}_g{n}"] = pv_profile.values
            
            elif df_Gen_Renewable.loc[n, "Renewable Type"] == "Wind":
                grid.add("Generator", f"Wind{location}_g{n}", #g{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                    bus = f"Bus_node_{location}", 
                    p_nom = df_Gen_Renewable.loc[n, "Rated active power (MW)"],
                    p_min_pu = 0, #No hay restricción de potencia mínima para las renovables
                    marginal_cost = 0, #El coste marginal para las renovables se considera nulo
                    carrier= "AC")
                grid.generators_t.p_max_pu[f"Wind{location}_g{n}"] = wind_profile.values

def solve_opf(grid: pypsa.Network, solver_name) -> None:
    grid.optimize(solver_name=solver_name)

def plot_dispatch_figure_weekly_average(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del balance energético diario apilado
    en MWh/día a partir de una serie horaria en MW.
    """

    if horizon != "Multiperiod":
        return None

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Nos quedamos solo con las columnas relevantes
    cols_to_plot = pos_cols + neg_cols
    dispatch_week = dispatch_clean[cols_to_plot].resample("W").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_week.index)
    for col in pos_cols:
        y = dispatch_week[col]
        ax.fill_between(
            dispatch_week.index,
            base_pos,
            base_pos + y,
            step="mid",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_week.index,
            base_pos + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_week.index)
    for col in neg_cols:
        y = dispatch_week[col]
        ax.fill_between(
            dispatch_week.index,
            base_neg,
            base_neg + y,
            step="mid",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_week.index,
            base_neg + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Weekly energy balance")
    ax.set_ylabel("Energy [MWh/week]")
    ax.set_xlabel("Time")

    # Formato eje X para datos diarios
    n_weeks = len(dispatch_week)
    print(n_weeks)
    
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=int(13/127*n_weeks-136/127)+1)) #Función para autoajustar el número de ticks del eje x según el número de datos
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_dispatch_figure_daily_average(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del balance energético diario apilado
    en MWh/día a partir de una serie horaria en MW.
    """

    if horizon != "Multiperiod":
        return None

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Nos quedamos solo con las columnas relevantes
    cols_to_plot = pos_cols + neg_cols
    dispatch_daily = dispatch_clean[cols_to_plot].resample("D").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_daily.index)
    for col in pos_cols:
        y = dispatch_daily[col]
        ax.fill_between(
            dispatch_daily.index,
            base_pos,
            base_pos + y,
            step="mid",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_daily.index,
            base_pos + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_daily.index)
    for col in neg_cols:
        y = dispatch_daily[col]
        ax.fill_between(
            dispatch_daily.index,
            base_neg,
            base_neg + y,
            step="mid",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_daily.index,
            base_neg + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Daily energy balance")
    ax.set_ylabel("Energy [MWh/day]")
    ax.set_xlabel("Time")

    # Formato eje X para datos diarios
    n_days = len(dispatch_daily)
    print(n_days)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=int(7/130*n_days)))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_dispatch_figure_hourly_snapshots(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del dispatch apilado.
    """

    if horizon != "Multiperiod":
        return None

    fig, ax = plt.subplots(figsize=(14, 6))

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_clean.index)
    for col in pos_cols:
        y = dispatch_clean[col]
        ax.fill_between(
            dispatch_clean.index,
            base_pos,
            base_pos + y,
            step="post",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_clean.index,
            base_pos + y,
            where="post",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_clean.index)
    for col in neg_cols:
        y = dispatch_clean[col]
        ax.fill_between(
            dispatch_clean.index,
            base_neg,
            base_neg + y,
            step="post",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Dispatch")
    ax.set_ylabel("Power [MW]")
    ax.set_xlabel("Time")

    # Formato eje X
    n_snapshots = len(dispatch_clean)
    print(n_snapshots)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=int(1/15*n_snapshots-1/5)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Hh\n%d %b"))

    else:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=int(1/408*n_snapshots+9/17)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_total_soc_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total de todas las baterías.
    """

    if horizon != "Multiperiod":
        return None

    fig, ax = plt.subplots(figsize=(10, 4))

    soc = grid.storage_units_t.state_of_charge.copy()
    soc_total = soc.sum(axis=1)

    ax.plot(soc_total.index, soc_total.values, label="Total SOC")

    capacity = (grid.storage_units["p_nom"] * grid.storage_units["max_hours"]).sum()
    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC")

    n_snapshots = len(soc_total)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_soc_per_battery_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC separado por baterías.
    """

    if horizon != "Multiperiod":
        return None

    soc = grid.storage_units_t.state_of_charge.copy()

    if soc.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    for battery_name in soc.columns:
        ax.plot(soc.index, soc[battery_name], label=battery_name)

    # Línea de capacidad máxima por batería si existe
    for battery_name in soc.columns:
        if battery_name in grid.storage_units.index:
            p_nom = grid.storage_units.loc[battery_name, "p_nom"]
            max_hours = grid.storage_units.loc[battery_name, "max_hours"]
            capacity = p_nom * max_hours
            ax.axhline(
                y=capacity,
                linestyle="--",
                linewidth=1,
                alpha=0.5
            )

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Battery SOC by unit")

    n_snapshots = len(soc)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_soc_per_battery_daily_average_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC medio diario separado por baterías.
    """

    if horizon != "Multiperiod":
        return None

    soc = grid.storage_units_t.state_of_charge.copy()

    if soc.empty:
        return None

    # Media diaria por batería
    soc_daily = soc.resample("D").mean()

    fig, ax = plt.subplots(figsize=(12, 5))

    for battery_name in soc_daily.columns:
        ax.plot(soc_daily.index, soc_daily[battery_name], label=battery_name)

    # Línea de capacidad máxima por batería si existe
    for battery_name in soc_daily.columns:
        if battery_name in grid.storage_units.index:
            p_nom = grid.storage_units.loc[battery_name, "p_nom"]
            max_hours = grid.storage_units.loc[battery_name, "max_hours"]
            capacity = p_nom * max_hours
            ax.axhline(
                y=capacity,
                linestyle="--",
                linewidth=1,
                alpha=0.5
            )

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Battery SOC by unit (daily average)")

    n_days = len(soc_daily)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_days / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_total_soc_daily_stats_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total con:
    - media diaria
    - mínimo diario
    - máximo diario
    - banda sombreada min-max
    """

    if horizon != "Multiperiod":
        return None

    soc = grid.storage_units_t.state_of_charge.copy()
    soc_total = soc.sum(axis=1)

    soc_daily_mean = soc_total.resample("D").mean()
    soc_daily_min = soc_total.resample("D").min()
    soc_daily_max = soc_total.resample("D").max()

    fig, ax = plt.subplots(figsize=(10, 4))

    # Banda min-max
    ax.fill_between(
        soc_daily_mean.index,
        soc_daily_min.values,
        soc_daily_max.values,
        alpha=0.25,
        label="Daily min-max range"
    )

    # Curvas
    ax.plot(soc_daily_mean.index, soc_daily_mean.values, linewidth=2, label="Daily mean SOC")
    ax.plot(soc_daily_min.index, soc_daily_min.values, linestyle="--", linewidth=1.2, label="Daily minimum SOC")
    ax.plot(soc_daily_max.index, soc_daily_max.values, linestyle="--", linewidth=1.2, label="Daily maximum SOC")

    capacity = (grid.storage_units["p_nom"] * grid.storage_units["max_hours"]).sum()
    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC (daily statistics)")

    n_days = len(soc_daily_mean)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_days / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_total_soc_weekly_stats_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total con:
    - media semanal
    - mínimo semanal
    - máximo semanal
    - banda sombreada min-max
    """

    if horizon != "Multiperiod":
        return None

    soc = grid.storage_units_t.state_of_charge.copy()
    soc_total = soc.sum(axis=1)

    soc_weekly_mean = soc_total.resample("W").mean()
    soc_weekly_min = soc_total.resample("W").min()
    soc_weekly_max = soc_total.resample("W").max()

    fig, ax = plt.subplots(figsize=(10, 4))

    # Banda min-max
    ax.fill_between(
        soc_weekly_mean.index,
        soc_weekly_min.values,
        soc_weekly_max.values,
        alpha=0.25,
        label="Weekly min-max range"
    )

    # Curvas
    ax.plot(soc_weekly_mean.index, soc_weekly_mean.values, linewidth=2, label="Weekly mean SOC")
    ax.plot(soc_weekly_min.index, soc_weekly_min.values, linestyle="--", linewidth=1.2, label="Weekly minimum SOC")
    ax.plot(soc_weekly_max.index, soc_weekly_max.values, linestyle="--", linewidth=1.2, label="Weekly maximum SOC")

    capacity = (grid.storage_units["p_nom"] * grid.storage_units["max_hours"]).sum()
    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC (weekly statistics)")

    n_weeks = len(soc_weekly_mean)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_energy_balance_sankey(dispatch_clean, grid):
    """
    Plot Sankey diagram for the aggregated microgrid energy balance.

    Parameters
    ----------
    dispatch_clean : pd.DataFrame
        DataFrame with aggregated dispatch results.
    grid : pypsa.Network
        PyPSA network object.
    """

    def safe_sum(df, col):
        return df[col].sum() if col in df.columns else 0.0

    pv = safe_sum(dispatch_clean, "PV")
    wind = safe_sum(dispatch_clean, "Wind")
    grid_import = safe_sum(dispatch_clean, "Grid_import")
    battery_discharge = safe_sum(dispatch_clean, "battery_discharge")
    dispatch = safe_sum(dispatch_clean, "Dispatch")

    load = grid.loads_t.p.sum().sum() if hasattr(grid.loads_t.p, "sum") else 0.0
    battery_charge = -safe_sum(dispatch_clean, "battery_charge")
    grid_export = -safe_sum(dispatch_clean, "Grid_export")

    total_in = pv + wind + grid_import + battery_discharge + dispatch
    total_out = load + battery_charge + grid_export

    print("Total input:", total_in)
    print("Total output:", total_out)

    labels = [
        "PV",                 # 0
        "Wind",               # 1
        "Grid import",        # 2
        "Battery discharge",  # 3
        "Dispatch",           # 4
        "Energy supplied",    # 5
        "Load",               # 6
        "Battery charge",     # 7
        "Grid export"         # 8
    ]

    source = [
        0, 1, 2, 3, 4,
        5, 5, 5
    ]

    target = [
        5, 5, 5, 5, 5,
        6, 7, 8
    ]

    value = [
        pv, wind, grid_import, battery_discharge, dispatch,
        load, battery_charge, grid_export
    ]

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=25,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=labels
        ),
        link=dict(
            source=source,
            target=target,
            value=value
        )
    )])

    fig.update_layout(
        title=dict(
            text="Microgrid energy balance",
            x=0.03,
            y=0.95
        ),
        font=dict(size=16),
        height=750,
        margin=dict(l=20, r=20, t=100, b=20)
    )

    return fig

def dispatch_graph_resolution_choice(df_SYS_settings: pd.DataFrame, dispatch_clean: pd.DataFrame)-> None:
    if df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]=="Auto":
        if df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]>=200:
            plot_dispatch_figure_weekly_average(dispatch_clean)
        elif 60<=df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]<200:
            plot_dispatch_figure_daily_average(dispatch_clean)
        else:
            plot_dispatch_figure_hourly_snapshots(dispatch_clean)
    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]=="Hourly":
        plot_dispatch_figure_hourly_snapshots(dispatch_clean)
    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]=="Daily":
        plot_dispatch_figure_daily_average(dispatch_clean)
    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]=="Weekly":
        plot_dispatch_figure_weekly_average(dispatch_clean)

def export_results(grid: pypsa.Network, df_SYS_settings: pd.DataFrame)-> None:
    # Agrupamos todos los pwl de un mismo generador
    cols_base = grid.generators_t.p.columns.str.replace(r'_seg\d+$', '', regex=True)
    dispatch = grid.generators_t.p.T.groupby(cols_base).sum().T

    dispatch["PV"] = dispatch[[c for c in dispatch.columns if "PV" in c]].sum(axis=1) #Agrupamos toda la generación fotovoltaica
    dispatch["Wind"] = dispatch[[c for c in dispatch.columns if "Wind" in c]].sum(axis=1) #Agrupamos toda la generación eólica 
    dispatch["Dispatch"] = dispatch[[c for c in dispatch.columns if "Dispatch" in c]].sum(axis=1) #Agrupamos todos los generadores despachables

    battery_discharge = grid.storage_units_t.p.clip(lower=0).sum(axis=1) #Agrupamos las descargas de todas las baterías
    battery_charge = grid.storage_units_t.p.clip(upper=0).sum(axis=1) #Agrupamos las cargas de todas las baterías
    
    dispatch.insert(0, "battery_discharge", battery_discharge) #Incluimos en el dataframe del despacho la descarga de las baterías
    dispatch.insert(1, "battery_charge", battery_charge) #Incluimos en el dataframe del despacho la carga de las baterías
    dispatch["Grid_export"] = -dispatch["Grid_export"] #La exportación de energía a la red la tomamos como negativa


    # Nos quedamos solo con las columnas agregadas que queremos mostrar
    dispatch_clean = pd.DataFrame(index=dispatch.index)
    dispatch_clean["PV"] = dispatch["PV"]
    dispatch_clean["Wind"] = dispatch["Wind"]
    dispatch_clean["battery_discharge"] = dispatch["battery_discharge"]
    dispatch_clean["Dispatch"] = dispatch["Dispatch"]
    dispatch_clean["Grid_import"] = dispatch["Grid_import"]
    dispatch_clean["battery_charge"] = dispatch["battery_charge"]
    dispatch_clean["Grid_export"] = dispatch["Grid_export"]

    # Eliminamos columnas que sean todo ceros o casi todo ceros
    dispatch_clean = dispatch_clean.loc[:, (dispatch_clean.abs() > 1e-6).any()]
    
    # GRÁFICO DE DESPACHO ESCALONADO SOLO PARA OPF MULTIPERIODO

    #dispatch_graph_resolution_choice(df_SYS_settings, dispatch_clean)
    # GRÁFICO DE SOC DE TODAS LAS BATERÍAS SUMADAS
    plot_total_soc_figure(grid)
    plot_total_soc_daily_stats_figure(grid)
    plot_total_soc_weekly_stats_figure(grid)
    # GRÁFICO DE SOC DE CADA BATERÍA
    plot_soc_per_battery_daily_average_figure(grid)
    plt.show()
    # GRÁFICO DE DESPACHO TIPO SANKEY 
    #fig3 = plot_energy_balance_sankey(dispatch_clean, grid)
    #fig3.show()
    # EXPORTACIÓN A EXCEL
    with pd.ExcelWriter("results.xlsx", engine="openpyxl") as writer:
        dispatch_clean.round(2).to_excel(writer, sheet_name="dispatch")
        grid.storage_units_t.p.round(2).to_excel(writer, sheet_name="battery_power")
        grid.storage_units_t.state_of_charge.round(2).to_excel(writer, sheet_name="battery_soc")
        grid.lines_t.p0.round(2).to_excel(writer, sheet_name="line_flows")
        grid.buses_t.marginal_price.round(2).to_excel(writer, sheet_name="prices")
   
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
    df_Grid_connection = data["Grid_connection"]

    df_TS_Wind_Profiles = data["TS_Wind_Profiles"]
    df_TS_Wind_Profiles["time"] = pd.to_datetime(df_TS_Wind_Profiles["time"]).dt.tz_localize(None)
    df_TS_Wind_Profiles = df_TS_Wind_Profiles.set_index("time")

    df_TS_PV_Profiles = data["TS_PV_Profiles"]
    df_TS_PV_Profiles["time"] = pd.to_datetime(df_TS_PV_Profiles["time"]).dt.tz_localize(None)
    df_TS_PV_Profiles = df_TS_PV_Profiles.set_index("time")

    df_TS_Energy_Prices = data["TS_Energy_Prices"]
    df_TS_Energy_Prices["time"] = pd.to_datetime(df_TS_Energy_Prices["time"]).dt.tz_localize(None)
    df_TS_Energy_Prices = df_TS_Energy_Prices.set_index("time")
    
    df_TS_LoadProfiles = data["TS_LoadProfiles"]
    df_TS_LoadProfiles["time"] = pd.to_datetime(df_TS_LoadProfiles["time"]).dt.tz_localize(None)
    df_TS_LoadProfiles = df_TS_LoadProfiles.set_index("time")


    grid = build_network(df_SYS_settings)

    add_buses(grid, df_Net_Buses)
    add_lines(grid, df_Net_Lines)
    add_loads(grid, df_Net_Loads, df_SYS_settings, df_TS_LoadProfiles) #Puesto que incluimos los generadores VOLL allí donde haya cargas 
    #la función add_loads debe recibir también df_SYS_settings para leer el VOLL en €/MWh introducido por el usuario

    add_dispatchable_generators(grid, df_Gen_Dispatchable)
    add_renewable_generator(grid, df_Gen_Renewable, df_SYS_settings, df_TS_Wind_Profiles, df_TS_PV_Profiles)
    add_storage_unit(grid, df_StorageUnit)
    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)


    solver=str(df_SYS_settings.loc[2, "SYSTEM PARAMETERS"])
    solve_opf(grid, solver_name=solver)

    #print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])
    #print(grid.storage_units[["p_nom", "max_hours", "bus", "efficiency_store", "efficiency_dispatch", "standing_loss", "state_of_charge_initial", "cyclic_state_of_charge", "marginal_cost"]])
    #print("\n")
    #print(grid.generators_t.p)

    #print(grid.objective)
    
    export_results(grid, df_SYS_settings)




if __name__ == "__main__":
    main()




