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

from Postprocessing.export_multiperiod_results import export_multiperiod_results
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
        "mip_rel_gap",
        "time_limit",
    ]

    # Validar básicos
    missing_base = [k for k in base_required if k not in gui_params]
    if missing_base:
        raise KeyError(f"Faltan parámetros básicos de la GUI: {missing_base}")

    horizon = gui_params["Static / Multiperiod"]

    data = {
        "VOLL (€/MWh)": float(gui_params["VOLL (€/MWh)"]),
        "Static / Multiperiod": horizon,
        "mip_rel_gap": float(gui_params.get("mip_rel_gap", 0.001)),
        "time_limit": int(gui_params.get("time_limit", 3600)),
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


def solve_opf(
    grid: pypsa.Network,
    solver_name: str,
    battery_specs=None,
    mip_rel_gap: float | None = None,
    time_limit: int | None = None,
) -> tuple:
    
    if battery_specs is not None:
        extra_func = lambda n, sns: add_battery_constraints(n, sns, battery_specs)
    else:
        extra_func = None

    solver_options = {}

    if mip_rel_gap is not None:
        solver_options["mip_rel_gap"] = mip_rel_gap

    if time_limit is not None:
        solver_options["time_limit"] = time_limit

    status, condition = grid.optimize(
        solver_name=solver_name,
        extra_functionality=extra_func,
        solver_options=solver_options,
    )

    print("Optimization status:", status)
    print("Optimization condition:", condition)

    return status, condition

def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_default_input_file() -> Path:
    return get_base_dir() / "GridInputs.xlsx"


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
  

    df_SYS_settings = build_sys_settings_from_gui(system_parameters)
    report_progress(20, "parámetros de la GUI cargados")
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    startdate = params["Start date (dd/mm/aaaa)"]
    duration = params["Simulation duration (days)"]

    if horizon == "Static":
        static_snapshot_datetime = params["Static snapshot datetime"]
        startdate = static_snapshot_datetime.date()
        print(startdate)
        duration = 1
        
    

    # -- TRATAMIENTO DE PERFILES DE DEMANDA de ESPAÑA -- 

    BASE_DIR = Path(__file__).resolve().parent
    ruta1 = BASE_DIR / "System_data" / "ConsumosAnualesX_CCAA.csv"
    df_consumos_anuales_CCAA = pd.read_csv(ruta1, sep=";", usecols=range(1,5))
    df_consumos_anuales_CCAA = df_consumos_anuales_CCAA[df_consumos_anuales_CCAA["Producto consumido"]=="Electricidad"]
    df_consumos_anuales_CCAA = df_consumos_anuales_CCAA.dropna()

    demr_folder = BASE_DIR / "System_data" / "DemandaDiariaSistemaEléctricoPeninsular"
    df_demanda_ccaa = build_hourly_demand_by_region(
        df_consumos_anuales_CCAA=df_consumos_anuales_CCAA,
        demr_folder=demr_folder,
        startdate=startdate,
        days=duration,
        annual_unit="ktep"
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
    ruta = Path("GridInputs.xlsx").resolve()
    # Lectura de potencia instalada
    df_solar_instaled_capacity = pd.read_excel(ruta, sheet_name="Gen_PV_and_Wind", usecols="G:AB", skiprows=2, nrows=14 )
    df_wind_instaled_capacity = pd.read_excel(ruta, sheet_name="Gen_PV_and_Wind", usecols="G:AB", skiprows=19, nrows=14 )
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
    print(df_solar_node_profiles)
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


    # -- CALCULO DE LAS SERIES TEMPORALES DE INFLOW PARA GENERADORES HIDRÁULICOS SIN BOMBEO ESCALADO A PARTIR DEL RUNOFF --
    df_hydro_inflow_scaled = build_hydro_inflow(
    base_dir=BASE_DIR,
    startdate=startdate,
    days=duration)

    # -- SERIE TEMPORAL DE P_MAX_PU DE GENERADORES ROR ESCALADA A PARTIR DEL RUNOFF --
    df_ror_p_max_pu_scaled = build_ror_p_max_pu(base_dir=BASE_DIR, startdate=startdate, days=duration)

    # -- CALCULO DEL PORCENTAJE INICIAL DE LLENADO DE LOS EMBALSES HYDRO
    BASE_DIR = Path(__file__).resolve().parent

    ruta_llenado_embalses = BASE_DIR / "System_data" / "Embalses.xlsx"

    df_embalses = pd.read_excel(ruta_llenado_embalses, usecols="D:G")

    df_embalses = df_embalses[df_embalses["ELECTRICO_FLAG"]==1]

    df_embalses = get_embalses_closest_date(
    df=df_embalses,
    target_date=startdate
    )

    initial_soc_fraction = (
    df_embalses["AGUA_ACTUAL"].sum()
    / df_embalses["AGUA_TOTAL"].sum()
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
    report_progress(52, "cargas añadidas")
    add_dispatchable_generators(grid, df_Gen_Dispatchable, gas_price, co2_price, df_ror_p_max_pu_scaled)
    report_progress(58, "generadores despachables añadidos")
    add_renewable_generator(grid, params, df_solar_node_profiles, df_wind_node_profiles)

    report_progress(64, "generadores renovables añadidos")
    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)
    report_progress(70, "conexión a red añadida")

    if grid.buses.x.isna().any() or grid.buses.y.isna().any():
        print("There are buses for which latitude/longitude were not specified, therefore the grid will not be drawn")
    else:
        drawrealgrid(grid, df_Net_Buses, "Iberian_penninsula_grid.png")


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



    report_progress(80, "optimizando")

    mip_rel_gap = params["mip_rel_gap"]
    time_limit = params["time_limit"]

    status, condition = solve_opf(
    grid=grid,
    solver_name="highs",
    battery_specs=battery_specs,
    mip_rel_gap=mip_rel_gap,
    time_limit=time_limit,
    )
    
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

        export_multiperiod_results(grid, df_SYS_settings, df_available_renewable, CFsolar, CFwind)
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
