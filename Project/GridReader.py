from __future__ import annotations
import pandas as pd
import pypsa
from Results.export_multiperiod_results import export_multiperiod_results
from Results.export_static_results import export_static_results
from Network.build_network import build_network
from Network.buses import add_buses
from Network.grid_connection import *
from Network.lines import add_lines
from Network.loads import *
from Generators.renewable import *
from Generators.dispatchable import add_dispatchable_generators
from Storage.constraints import add_battery_constraints
from Storage.storage_model import add_storage_as_store_links


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

def solve_opf(grid: pypsa.Network, solver_name, battery_specs) -> None:
    grid.optimize(
    solver_name=solver_name,  
    extra_functionality=lambda n, sns: add_battery_constraints(n, sns, battery_specs)
)
    


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
    
    df_available_renewable = add_renewable_generator(grid, 
                                    df_Gen_Renewable, 
                                    df_SYS_settings, df_TS_Wind_Profiles, df_TS_PV_Profiles)
    
    grid_connection(grid, df_Grid_connection, df_TS_Energy_Prices, df_SYS_settings)


    solver=str(df_SYS_settings.loc[2, "SYSTEM PARAMETERS"])
    battery_specs = add_storage_as_store_links(grid, df_StorageUnit)
    solve_opf(grid, solver, battery_specs)
    #print(grid.generators[["p_nom", "p_min_pu", "marginal_cost", "bus"]])
    #print(grid.storage_units[["p_nom", "max_hours", "bus", "efficiency_store", "efficiency_dispatch", "standing_loss", "state_of_charge_initial", "cyclic_state_of_charge", "marginal_cost"]])
    #print("\n")
    #print(grid.generators_t.p)

    #print(grid.objective)

    #print(grid.lines.loc[:, "s_nom"])

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    if horizon == "Multiperiod":
        export_multiperiod_results(grid, df_SYS_settings, df_available_renewable)
    elif horizon == "Static":
        export_static_results(grid)




if __name__ == "__main__":
    main()






