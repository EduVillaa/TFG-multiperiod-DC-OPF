import pypsa
import pandas as pd


def wind_series_reader(
    df_SYS_settings: pd.DataFrame,
    df_TS_Wind_Profiles: pd.DataFrame,
    region: str
) -> pd.Series:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon == "Static":
        return pd.Series([1.0], name=region)

    elif horizon == "Multiperiod":
        start_date = params["Start date (dd/mm/aaaa)"]
        simulation_days = params["Simulation duration (days)"]

        if pd.isna(start_date):
            raise ValueError("Falta 'Start date (dd/mm/aaaa)' para horizonte Multiperiod.")
        if pd.isna(simulation_days):
            raise ValueError("Falta 'Simulation duration (days)' para horizonte Multiperiod.")

        simulation_hours = int(simulation_days) * 24
        return df_TS_Wind_Profiles.loc[start_date:, region].iloc[:simulation_hours]

    else:
        raise ValueError(f"Horizonte no reconocido: {horizon}")


def pv_series_reader(
    df_SYS_settings: pd.DataFrame,
    df_TS_PV_Profiles: pd.DataFrame,
    region: str
) -> pd.Series:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon == "Static":
        return pd.Series([1.0], name=region)

    elif horizon == "Multiperiod":
        start_date = params["Start date (dd/mm/aaaa)"]
        simulation_days = params["Simulation duration (days)"]

        if pd.isna(start_date):
            raise ValueError("Falta 'Start date (dd/mm/aaaa)' para horizonte Multiperiod.")
        if pd.isna(simulation_days):
            raise ValueError("Falta 'Simulation duration (days)' para horizonte Multiperiod.")

        simulation_hours = int(simulation_days) * 24
        return df_TS_PV_Profiles.loc[start_date:, region].iloc[:simulation_hours]

    else:
        raise ValueError(f"Horizonte no reconocido: {horizon}")

def add_renewable_generator(
    grid: pypsa.Network,
    df_Gen_Renewable: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
    df_TS_Wind_Profiles: pd.DataFrame,
    df_TS_PV_Profiles: pd.DataFrame
)-> None:

    for n in range(df_Gen_Renewable["GENERATOR LOCATION"].count()):
        location = int(df_Gen_Renewable.loc[n, "GENERATOR LOCATION"])

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
                    carrier="PV"
                )
                region = df_Gen_Renewable.loc[n, "Region"]
                pv_profile = pv_series_reader(df_SYS_settings, df_TS_PV_Profiles, region)
                grid.generators_t.p_max_pu[gen_name] = pv_profile.values

            elif df_Gen_Renewable.loc[n, "Renewable Type"] == "Wind":
                gen_name = f"Wind{location}_{n}"

                grid.add(
                    "Generator",
                    gen_name,
                    bus=f"Bus_node_{location}",
                    p_nom=p_nom,
                    p_min_pu=0,
                    marginal_cost=0,
                    carrier="Wind"
                )

                region = df_Gen_Renewable.loc[n, "Region"]
                wind_profile = wind_series_reader(df_SYS_settings, df_TS_Wind_Profiles, region)
                grid.generators_t.p_max_pu[gen_name] = wind_profile.values


def build_available_renewable_df(
    df_Gen_Renewable: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
    df_TS_Wind_Profiles: pd.DataFrame,
    df_TS_PV_Profiles: pd.DataFrame
) -> pd.DataFrame:
    """
    Devuelve un dataframe con la potencia renovable disponible [MW]
    de cada generador renovable, teniendo en cuenta la región específica
    de cada uno.
    """

    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]

    if horizon == "Static":
        index = pd.RangeIndex(1)

    elif horizon == "Multiperiod":
        start_date = params["Start date (dd/mm/aaaa)"]
        simulation_days = params["Simulation duration (days)"]

        if pd.isna(start_date):
            raise ValueError("Falta 'Start date (dd/mm/aaaa)' para horizonte Multiperiod.")
        if pd.isna(simulation_days):
            raise ValueError("Falta 'Simulation duration (days)' para horizonte Multiperiod.")

        simulation_hours = int(simulation_days) * 24
        index = df_TS_Wind_Profiles.loc[start_date:].iloc[:simulation_hours].index

    else:
        raise ValueError(f"Horizonte no reconocido: {horizon}")

    df_available_renewable = pd.DataFrame(index=index)

    for n in range(df_Gen_Renewable["GENERATOR LOCATION"].count()):
        location = int(df_Gen_Renewable.loc[n, "GENERATOR LOCATION"])

        if pd.isna(location):
            continue

        region = df_Gen_Renewable.loc[n, "Region"]
        p_nom = df_Gen_Renewable.loc[n, "Rated active power (MW)"]
        tech = df_Gen_Renewable.loc[n, "Renewable Type"]

        if pd.isna(region) or pd.isna(p_nom) or pd.isna(tech):
            continue

        if tech == "PV":
            gen_name = f"PV{location}_{n}"
            pv_profile = pv_series_reader(df_SYS_settings, df_TS_PV_Profiles, region)
            df_available_renewable[gen_name] = pv_profile.values * p_nom

        elif tech == "Wind":
            gen_name = f"Wind{location}_{n}"
            wind_profile = wind_series_reader(df_SYS_settings, df_TS_Wind_Profiles, region)
            df_available_renewable[gen_name] = wind_profile.values * p_nom

    return df_available_renewable
   