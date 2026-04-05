import pypsa
import pandas as pd


def add_loads(grid: pypsa.Network, df_Net_Loads: pd.DataFrame, df_SYS_settings: pd.DataFrame, df_TS_LoadProfiles: pd.DataFrame) -> None:
    params = df_SYS_settings["SYSTEM PARAMETERS"]

    for n in range(df_Net_Loads["Active power demand (MW)"].last_valid_index() + 1):
        
        if df_Net_Loads.loc[n, "Time series load profile"] == "Residential (P2.0TD)":
            load_profile_type = "COEF. PERFIL P2.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Commercial / Industrial (P3.0TD)":
            load_profile_type = "COEF. PERFIL P3.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Electric vehicle (P3.0TDVE)":
            load_profile_type = "COEF. PERFIL P3.0TDVE"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Flat load profile":
            load_profile_type = "Flat LP"

        AnnualConsumption = df_Net_Loads.loc[n, "Annual energy consumption (MWh/year)"]
        load_profile_MW = load_profile_reader(df_TS_LoadProfiles, df_SYS_settings, load_profile_type) * AnnualConsumption
        horizon = params["Static / Multiperiod"]
        location = df_Net_Loads.loc[n, "LOAD LOCATION"]
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]
        VOLL = params["VOLL (€/MWh)"] # €/MWh (valor alto)
        if pd.notna(Pd):
            grid.add("Load", f"Load_node_{location}_L{n}_{load_profile_type}",  #L{n} permite distinguir cargas del mismo nodo
                    bus=f"Bus_node_{location}", 
                    p_set=Pd, 
                    carrier="AC")
            
            if horizon == "Multiperiod":
                grid.loads_t.p_set.loc[:, f"Load_node_{location}_L{n}_{load_profile_type}"] = load_profile_MW

            use_shed = True
            if use_shed:
                grid.add("Generator", f"shedding_gen_node_{location}", bus=f"Bus_node_{location}", 
                        p_nom=1e6, 
                        marginal_cost=VOLL,
                        p_min_pu=0,
                        carrier="AC")


def load_profile_reader(df_TS_LoadProfiles: pd.DataFrame, df_SYS_settings: pd.DataFrame, load_profile_type)-> pd.Series:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    start_Date = params["Start date (dd/mm/aaaa)"]
    horizon = params["Static / Multiperiod"]
    simulation_days = params["Simulation duration (days)"]
    simulation_hours = simulation_days*24

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        df = df_TS_LoadProfiles.loc[start_Date:].iloc[:simulation_hours]
        return df.loc[:, load_profile_type]
