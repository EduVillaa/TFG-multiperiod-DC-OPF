from __future__ import annotations

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
from Network_builder.Generators.renewable import *
from Network_builder.Generators.dispatchable import add_dispatchable_generators
from Network_builder.Storage.constraints import add_battery_constraints
from Network_builder.Storage.storage_model import add_storage_as_store_links
from Postprocessing.drawgridinmap import drawrealgrid

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

    sheets["Net_Loads"] = pd.read_excel(
        filename,
        sheet_name="Net_Loads",
        header=3
    ).iloc[:, 1:]

    sheets["Gen_Dispatchable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Dispatchable",
        header=2
    ).iloc[:, 1:]

    sheets["Gen_Renewable"] = pd.read_excel(
        filename,
        sheet_name="Gen_Renewable",
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

    sheets["TS_Wind_Profiles"] = pd.read_excel(
        filename,
        sheet_name="TS_Wind_Profiles",
        header=0
    ).iloc[:, 0:]

    sheets["TS_PV_Profiles"] = pd.read_excel(
        filename,
        sheet_name="TS_PV_Profiles",
        header=0
    ).iloc[:, 0:]

    sheets["TS_Energy_Prices"] = pd.read_excel(
        filename,
        sheet_name="TS_Energy_Prices",
        header=0
    ).iloc[:, 0:]

    sheets["TS_LoadProfiles"] = pd.read_excel(
        filename,
        sheet_name="TS_LoadProfiles",
        header=0
    ).iloc[:, 0:]

    return sheets


def build_sys_settings_from_gui(gui_params: dict) -> pd.DataFrame:
    # Parámetros siempre obligatorios
    base_required = [
        "VOLL (€/MWh)",
        "Static / Multiperiod",
    ]

    # Validar básicos
    missing_base = [k for k in base_required if k not in gui_params]
    if missing_base:
        raise KeyError(f"Faltan parámetros básicos de la GUI: {missing_base}")

    horizon = gui_params["Static / Multiperiod"]

    data = {
        "VOLL (€/MWh)": float(gui_params["VOLL (€/MWh)"]),
        "Static / Multiperiod": horizon,
    }

    # Solo si es multiperiod, añadimos el resto
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
        })

    else:
        # Opcional: puedes meter valores dummy o None
        data.update({
            "Start date (dd/mm/aaaa)": None,
            "Simulation duration (days)": None,
            "Graph resolution": None,
        })

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

def solve_opf(grid: pypsa.Network, solver_name: str, battery_specs=None) -> None:
    
    if battery_specs is not None:
        extra_func = lambda n, sns: add_battery_constraints(n, sns, battery_specs)
    else:
        extra_func = None

    grid.optimize(
        solver_name=solver_name,
        extra_functionality=extra_func
    )



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
    """
    Ejecuta el programa completo.
    Devuelve la ruta del archivo de entrada usado.
    Lanza excepción si algo falla.
    """
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
    print("Los datos de entrada han sido leidos")

    if system_parameters is None:
        raise ValueError("No se han recibido los parámetros del sistema desde la GUI.")
  

    df_SYS_settings = build_sys_settings_from_gui(system_parameters)
    report_progress(20, "parámetros de la GUI cargados")
    print("Se han obtenido los datos introducidos en la GUI")
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]


    economic_settings = build_battery_economic_settings_from_gui(system_parameters)
    
    if horizon=="Multiperiod":
        def crf(i, n):
            return (i * (1 + i)**n) / ((1 + i)**n - 1)
        
        interest = economic_settings.iloc[0, 0]/100
        lifetime = economic_settings.iloc[1, 0]
        CRF = crf(interest, lifetime)


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
    report_progress(30, "red creada")

    add_buses(grid, df_Net_Buses)
    report_progress(38, "buses añadidos")


    add_lines(grid, df_Net_Lines)
    report_progress(45, "líneas añadidas")
    

    if grid.buses.x.isna().any() or grid.buses.y.isna().any():
        print("There are buses for which latitude/longitude were not specified, therefore the grid will not be drawn")
    else:
        drawrealgrid(grid, df_Net_Buses, "azerbaijan_grid.png")
    print("drawrealgrid está OK")

    add_loads(grid, df_Net_Loads, df_SYS_settings, df_TS_LoadProfiles)
    report_progress(52, "cargas añadidas")
    print("error 2.1")
    add_dispatchable_generators(grid, df_Gen_Dispatchable)
    report_progress(58, "generadores despachables añadidos")
    print("error 2.2")
    add_renewable_generator(grid, df_Gen_Renewable, df_SYS_settings, df_TS_Wind_Profiles, df_TS_PV_Profiles)
    report_progress(64, "generadores renovables añadidos")
    print("error 2.3")
    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)
    report_progress(70, "conexión a red añadida")
    print("error 2.4")
    solver = "highs"
 
    if horizon == "Multiperiod":
        battery_specs = add_storage_as_store_links(df_SYS_settings, grid, df_StorageUnit, CRF)
    elif horizon == "Static":
        battery_specs=None
    report_progress(76, "almacenamiento configurado")
    print("error 2.5")

    report_progress(80, "optimizando")
    solve_opf(grid, solver, battery_specs)
    report_progress(90, "optimización terminada")
    
    print("ERROR 3")
    if horizon == "Multiperiod":
        df_available_renewable = build_available_renewable_df(df_Gen_Renewable, df_SYS_settings, df_TS_Wind_Profiles, df_TS_PV_Profiles)
        export_multiperiod_results(grid, df_SYS_settings, df_available_renewable)
    elif horizon == "Static":
        export_static_results(grid)
    else:
        raise ValueError(f"Valor de horizonte no reconocido: {horizon}")

    report_progress(100, "resultados exportados")
    print("ERROR 4")
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
