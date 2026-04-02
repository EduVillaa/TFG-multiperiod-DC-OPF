import pypsa
import pandas as pd

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

def add_renewable_generator(
    grid: pypsa.Network,
    df_Gen_Renewable: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
    df_TS_Wind_Profiles: pd.DataFrame,
    df_TS_PV_Profiles: pd.DataFrame
) -> pd.DataFrame:

    wind_profile = wind_series_reader(df_SYS_settings, df_TS_Wind_Profiles)
    pv_profile = pv_series_reader(df_SYS_settings, df_TS_PV_Profiles)

    # Índice temporal común
    df_available_renewable = pd.DataFrame(index=wind_profile.index)

    for n in range(df_Gen_Renewable["GENERATOR LOCATION"].count()):
        location = df_Gen_Renewable.loc[n, "GENERATOR LOCATION"]

        if pd.notna(location):
            p_nom = df_Gen_Renewable.loc[n, "Rated active power (MW)"]

            if df_Gen_Renewable.loc[n, "Renewable Type"] == "PV":
                gen_name = f"PV{location}_{n}"

                grid.add(
                    "Generator",
                    gen_name,
                    bus=f"Bus_node_{location}",
                    p_nom=p_nom,
                    p_min_pu=0,
                    marginal_cost=0,
                    carrier="AC"
                )

                grid.generators_t.p_max_pu[gen_name] = pv_profile.values

                # Potencia disponible [MW]
                df_available_renewable[gen_name] = pv_profile.values * p_nom

            elif df_Gen_Renewable.loc[n, "Renewable Type"] == "Wind":
                gen_name = f"Wind{location}_{n}"

                grid.add(
                    "Generator",
                    gen_name,
                    bus=f"Bus_node_{location}",
                    p_nom=p_nom,
                    p_min_pu=0,
                    marginal_cost=0,
                    carrier="AC"
                )

                grid.generators_t.p_max_pu[gen_name] = wind_profile.values

                # Potencia disponible [MW]
                df_available_renewable[gen_name] = wind_profile.values * p_nom

    return df_available_renewable
