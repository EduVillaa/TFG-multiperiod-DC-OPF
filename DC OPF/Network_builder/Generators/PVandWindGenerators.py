import pypsa
import pandas as pd


def add_renewable_generator(
    grid: pypsa.Network,
    params: pd.DataFrame,
    df_solar_node_profiles: pd.DataFrame,
    df_wind_node_profiles: pd.DataFrame
) -> None:

    # para evitar warnings de carriers no definidos
    for carrier in ["PV", "Wind"]:
        if carrier not in grid.carriers.index:
            grid.add("Carrier", carrier)

    horizon = params["Static / Multiperiod"]

    # -------------------------
    # Solar PV
    # -------------------------
    if horizon == "Static":
        df_solar_node_profiles = df_solar_node_profiles.copy()
        df_solar_node_profiles["time"] = pd.to_datetime(df_solar_node_profiles["time"])
        df_solar_node_profiles = df_solar_node_profiles.set_index("time")
        df_solar_node_profiles = df_solar_node_profiles.sort_index()
    
    for col in df_solar_node_profiles.columns[1:]:

        serie_mw = pd.to_numeric(df_solar_node_profiles[col], errors="coerce").fillna(0)
        p_nom = serie_mw.max()

        grid.add(
            "Generator",
            f"PV_{col}",
            bus=f"Bus.{col}",
            p_nom=p_nom,
            p_min_pu=0,
            marginal_cost=0,
            carrier="PV",
        )

        if horizon == "Multiperiod":
            if p_nom > 0:
                grid.generators_t.p_max_pu[f"PV_{col}"] = (serie_mw / p_nom).values
            else:
                grid.generators_t.p_max_pu[f"PV_{col}"] = 0.0
        elif horizon == "Static":

            snapshot = grid.snapshots[0]
            p_max_pu_series = (serie_mw / p_nom).copy()
            p_max_pu_series.index = pd.to_datetime(p_max_pu_series.index)
            p_max_pu_series = p_max_pu_series.sort_index()

            valor_estatico = p_max_pu_series.loc[snapshot]

            grid.generators_t.p_max_pu[f"PV_{col}"] = pd.Series(
                [valor_estatico],
                index=grid.snapshots
            )

    # -------------------------
    # Wind
    # -------------------------
    if horizon == "Static":
        df_wind_node_profiles = df_wind_node_profiles.copy()
        df_wind_node_profiles["time"] = pd.to_datetime(df_wind_node_profiles["time"])
        df_wind_node_profiles = df_wind_node_profiles.set_index("time")
        df_wind_node_profiles = df_wind_node_profiles.sort_index()

    for col in df_wind_node_profiles.columns[1:]:

        serie_mw = pd.to_numeric(df_wind_node_profiles[col], errors="coerce").fillna(0)
        p_nom = serie_mw.max()

        grid.add(
            "Generator",
            f"Wind_{col}",
            bus=f"Bus.{col}",
            p_nom=p_nom,
            p_min_pu=0,
            marginal_cost=0,
            carrier="Wind",
        )

        if horizon == "Multiperiod":
            if p_nom > 0:
                grid.generators_t.p_max_pu[f"Wind_{col}"] = (serie_mw / p_nom).values
            else:
                grid.generators_t.p_max_pu[f"Wind_{col}"] = 0.0
        elif horizon == "Static":
            snapshot = grid.snapshots[0]
            p_max_pu_series = (serie_mw / p_nom).copy()
            p_max_pu_series.index = pd.to_datetime(p_max_pu_series.index)
            p_max_pu_series = p_max_pu_series.sort_index()

            valor_estatico = p_max_pu_series.loc[snapshot]

            grid.generators_t.p_max_pu[f"Wind_{col}"] = pd.Series(
                [valor_estatico],
                index=grid.snapshots
            )
