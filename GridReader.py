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
    required_keys = [
        "VOLL (€/MWh)",
        "Static / Multiperiod",
        "Region",
        "Start date (dd/mm/aaaa)",
        "Simulation duration (days)",
        "Graph resolution",
    ]

    missing = [k for k in required_keys if k not in gui_params]
    if missing:
        raise KeyError(f"Faltan parámetros de la GUI: {missing}")

    df_SYS_settings = pd.DataFrame({
        "SYSTEM PARAMETERS": pd.Series({
            "VOLL (€/MWh)": float(gui_params["VOLL (€/MWh)"]),
            "Static / Multiperiod": gui_params["Static / Multiperiod"],
            "Region": gui_params["Region"],
            "Start date (dd/mm/aaaa)": pd.to_datetime(gui_params["Start date (dd/mm/aaaa)"]),
            "Simulation duration (days)": int(gui_params["Simulation duration (days)"]),
            "Graph resolution": gui_params["Graph resolution"],
        })
    })

    return df_SYS_settings

def solve_opf(grid: pypsa.Network, solver_name: str, battery_specs) -> None:
    grid.optimize(
        solver_name=solver_name,
        extra_functionality=lambda n, sns: add_battery_constraints(n, sns, battery_specs)
    )


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def get_default_input_file() -> Path:
    return get_base_dir() / "GridInputs.xlsx"


def run_program(
    input_file: str | Path | None = None,
    system_parameters: dict | None = None
) -> Path:
    """
    Ejecuta el programa completo.
    Devuelve la ruta del archivo de entrada usado.
    Lanza excepción si algo falla.
    """
    if input_file is None:
        input_path = get_default_input_file()
    else:
        input_path = Path(input_file).resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"No se encontró el archivo de entrada: {input_path}")

    data = leerhojas(input_path)

    if system_parameters is None:
        raise ValueError("No se han recibido los parámetros del sistema desde la GUI.")

    df_SYS_settings = build_sys_settings_from_gui(system_parameters)

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
    add_loads(grid, df_Net_Loads, df_SYS_settings, df_TS_LoadProfiles)
    add_dispatchable_generators(grid, df_Gen_Dispatchable)

    df_available_renewable = add_renewable_generator(
        grid,
        df_Gen_Renewable,
        df_SYS_settings,
        df_TS_Wind_Profiles,
        df_TS_PV_Profiles
    )

    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)

    solver = "highs"
    battery_specs = add_storage_as_store_links(grid, df_StorageUnit)
    solve_opf(grid, solver, battery_specs)

    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon == "Multiperiod":
        export_multiperiod_results(grid, df_SYS_settings, df_available_renewable)
    elif horizon == "Static":
        export_static_results(grid)
    else:
        raise ValueError(f"Valor de horizonte no reconocido: {horizon}")

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
