from __future__ import annotations
import pandas as pd
import pypsa
import matplotlib.pyplot as plt
import networkx as nx
from pathlib import Path
from typing import Optional, Dict, Any
import matplotlib.dates as mdates


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
    print(start_date)
    time_horizon = str(df_SYS_settings.loc[3, "SYSTEM PARAMETERS"])
    if time_horizon == "Static":
        grid.set_snapshots(pd.DatetimeIndex(["2026-01-01 00:00"]))

    elif time_horizon == "Day":
        grid.set_snapshots(pd.date_range(start_date, periods=24, freq="h"))

    elif time_horizon == "Week":
        grid.set_snapshots(pd.date_range(start_date, periods=168, freq="h"))
  
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

    if horizon == "Day":
        df = df_TS_LoadProfiles.loc[start_Date:].iloc[:24]
        return df.loc[:, load_profile_type]

    elif horizon == "Static":
        return pd.Series([1])

    elif horizon == "Week":
        df = df_TS_LoadProfiles.loc[start_Date:].iloc[:168]
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

        location = df_Net_Loads.loc[n, "LOAD LOCATION"]
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]
        Ploss = df_Net_Loads.loc[n, "Loss factor (%)"]
        VOLL = df_SYS_settings.loc[0, "SYSTEM PARAMETERS"] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{location}_L{n}",  #L{n} permite distinguir cargas del mismo nodo
                    bus=f"Bus_node_{location}", 
                    p_set=Pd*(1+Ploss), carrier="AC")
            
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

    if horizon == "Day":
        df = df_TS_Energy_Prices.loc[start_Date:].iloc[:24]
        return df["Precio mercado SPOT Diario España (€/MWh)"]

    elif horizon == "Static":
        return pd.Series([1])

    elif horizon == "Week":
        df = df_TS_Energy_Prices.loc[start_Date:].iloc[:168]
        return df["Precio mercado SPOT Diario España (€/MWh)"]

def wind_series_reader(df_SYS_settings: pd.DataFrame,
                            df_TS_Wind_Profiles: pd.DataFrame) -> pd.Series:
    
    region = str(df_SYS_settings.loc[4, "SYSTEM PARAMETERS"])
    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]

    if horizon == "Day":
        return df_TS_Wind_Profiles.loc[start_Date.strftime("%Y-%m-%d"), region]

    elif horizon == "Static":
        return pd.Series([1], name=region)

    elif horizon == "Week":
        return df_TS_Wind_Profiles.loc[start_Date:, region].iloc[:168]
    
def pv_series_reader(df_SYS_settings: pd.DataFrame,
                            df_TS_PV_Profiles: pd.DataFrame) -> pd.Series:
    region = str(df_SYS_settings.loc[4, "SYSTEM PARAMETERS"])
    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]

    if horizon == "Day":
        return df_TS_PV_Profiles.loc[start_Date.strftime("%Y-%m-%d"), region]

    elif horizon == "Static":
        return pd.Series([1], name=region)

    elif horizon == "Week":
        return df_TS_PV_Profiles.loc[start_Date:, region].iloc[:168]

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

def export_results(grid: pypsa.Network)-> None:
    # Agrupamos todos los pwl de un mismo generador
    cols_base = grid.generators_t.p.columns.str.replace(r'_seg\d+$', '', regex=True)
    dispatch = grid.generators_t.p.T.groupby(cols_base).sum().T

    dispatch["PV"] = dispatch[[c for c in dispatch.columns if "PV" in c]].sum(axis=1) #Agrupamos toda la generación fotovoltaica
    dispatch["Wind"] = dispatch[[c for c in dispatch.columns if "Wind" in c]].sum(axis=1) #Agrupamos toda la generación eólica 
    dispatch["Dispatch"] = dispatch[[c for c in dispatch.columns if "Dispatch" in c]].sum(axis=1) #Agrupamos todos los generadores despachables

    battery_discharge = grid.storage_units_t.p.clip(lower=0).sum(axis=1) #Agrupamos las descargas de todas las baterías
    battery_charge = grid.storage_units_t.p.clip(upper=0).sum(axis=1) #Agrupamos las cargas de todas las baterías

    print(grid.storage_units_t.p)
    
    dispatch.insert(0, "battery_discharge", battery_discharge) #Incluimos en el dataframe del despacho la descarga de las baterías
    dispatch.insert(1, "battery_charge", battery_charge) #Incluimos en el dataframe del despacho la carga de las baterías
    dispatch["Grid_export"] = -dispatch["Grid_export"] #La exportación de energía a la red la tomamos como negativa

    print(dispatch)


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

    # -----------------------------
    # GRÁFICO DE DESPACHO ESCALONADO
    # -----------------------------
    fig, ax = plt.subplots(figsize=(14, 6))

    # Orden recomendado
    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    # Dejamos solo las que existan realmente
    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    # Colores
    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Apilado positivo
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
        base_pos = base_pos + y

    # Apilado negativo
    base_neg = pd.Series(0.0, index=dispatch_clean.index)
    for col in neg_cols:
        y = dispatch_clean[col]   # ya es negativo
        ax.fill_between(
            dispatch_clean.index,
            base_neg,
            base_neg + y,
            step="post",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        base_neg = base_neg + y

    # Línea horizontal en cero
    ax.axhline(0, color="black", linewidth=1)

    # Título y etiquetas
    ax.set_title("Dispatch")
    ax.set_ylabel("Power [MW]")
    ax.set_xlabel("Time")

    # Formato del eje X
    # Para un día: cada 2 h
    # Para una semana: cada 12 h suele quedar bien
    n_snapshots = len(dispatch_clean)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))
    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
        ax.xaxis.set_minor_locator(mdates.HourLocator(interval=12))
        ax.xaxis.set_minor_formatter(mdates.DateFormatter("%Hh"))
        ax.tick_params(axis="x", which="major", pad=15)
        ax.tick_params(axis="x", which="minor", pad=3)
    else:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    # Leyenda
    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    plt.subplots_adjust(right=0.8, bottom=0.18)
    plt.tight_layout()

    # Guardamos figura
    #plt.savefig("dispatch_plot.png", dpi=300, bbox_inches="tight")
    #plt.close()
    plt.show()

    # -----------------------------
    # EXPORTACIÓN A EXCEL
    # -----------------------------
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



    print(grid.loads_t.p_set.head(168))

    solver=str(df_SYS_settings.loc[2, "SYSTEM PARAMETERS"])
    solve_opf(grid, solver_name=solver)

    #print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])
    #print(grid.storage_units[["p_nom", "max_hours", "bus", "efficiency_store", "efficiency_dispatch", "standing_loss", "state_of_charge_initial", "cyclic_state_of_charge", "marginal_cost"]])
    #print("\n")
    #print(grid.generators_t.p)

    #print(grid.objective)
    
    export_results(grid)

    """
    print(grid.loads_t.p)
    print(grid.links_t.p0)
    print(grid.links_t.p1)
    """




if __name__ == "__main__":
    main()




