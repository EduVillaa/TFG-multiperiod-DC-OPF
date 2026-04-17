import pypsa
import pandas as pd

def add_loads(
    grid: pypsa.Network,
    df_Net_Loads: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
    df_TS_LoadProfiles: pd.DataFrame
) -> None:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    VOLL = params["VOLL (€/MWh)"]

    for n in range(df_Net_Loads["Active power demand (MW)"].last_valid_index() + 1):

        if df_Net_Loads.loc[n, "Time series load profile"] == "Residential (P2.0TD)":
            load_profile_type = "COEF. PERFIL P2.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Commercial / Industrial (P3.0TD)":
            load_profile_type = "COEF. PERFIL P3.0TD"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Electric vehicle (P3.0TDVE)":
            load_profile_type = "COEF. PERFIL P3.0TDVE"
        elif df_Net_Loads.loc[n, "Time series load profile"] == "Flat load profile":
            load_profile_type = "Flat LP"
        else:
            raise ValueError(
                f"Perfil de carga no reconocido en la fila {n}: "
                f"{df_Net_Loads.loc[n, 'Time series load profile']}"
            )

        annual_consumption = df_Net_Loads.loc[n, "Annual energy consumption (MWh/year)"]
        location = df_Net_Loads.loc[n, "LOAD LOCATION"]
        Pd = df_Net_Loads.loc[n, "Active power demand (MW)"]

        if pd.notna(Pd):
            load_name = f"Load_node_{location}_L{n}_{load_profile_type}"

            grid.add(
                "Load",
                load_name,
                bus=f"Bus_node_{location}",
                p_set=Pd,
                carrier="AC"
            )

            if horizon == "Multiperiod":
                load_profile_MW = (
                    load_profile_reader(df_TS_LoadProfiles, df_SYS_settings, load_profile_type)
                    * annual_consumption
                )
                grid.loads_t.p_set.loc[:, load_name] = load_profile_MW.values

            use_shed = True
            if use_shed:
                grid.add(
                    "Generator",
                    f"shedding_gen_node_{location}",
                    bus=f"Bus_node_{location}",
                    p_nom=1e6,
                    marginal_cost=VOLL,
                    p_min_pu=0,
                    carrier="AC"
                )
                
def load_profile_reader(
    df_TS_LoadProfiles: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
    load_profile_type
) -> pd.Series:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon == "Static":
        return pd.Series([1.0])

    elif horizon == "Multiperiod":
        start_date = params["Start date (dd/mm/aaaa)"]
        simulation_days = params["Simulation duration (days)"]

        if pd.isna(start_date):
            raise ValueError("Falta 'Start date (dd/mm/aaaa)' para horizonte Multiperiod.")
        if pd.isna(simulation_days):
            raise ValueError("Falta 'Simulation duration (days)' para horizonte Multiperiod.")

        simulation_hours = int(simulation_days) * 24
        df = df_TS_LoadProfiles.loc[start_date:].iloc[:simulation_hours]
        return df.loc[:, load_profile_type]

    else:
        raise ValueError(f"Horizonte no reconocido: {horizon}")