import pypsa
import pandas as pd

def add_loads(
    grid: pypsa.Network,
    df_demanda_nodal: pd.DataFrame,
    df_SYS_settings: pd.DataFrame,
) -> None:
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    VOLL = params["VOLL (€/MWh)"]

    for col in df_demanda_nodal.columns[:]:

            grid.add(
                "Load",
                f"{col}_Load",
                bus=f"Bus.{col}",
                p_set=df_demanda_nodal[col].iloc[0],
                carrier="AC"
            )

            if horizon == "Multiperiod":
                load_profile_MW = (df_demanda_nodal[col]
                )
                grid.loads_t.p_set.loc[:, f"{col}_Load"] = load_profile_MW.values

            use_shed = True
            if use_shed:
                grid.add(
                    "Generator",
                    f"shedding_gen_node_{col}",
                    bus=f"Bus.{col}",
                    p_nom=1e6,
                    marginal_cost=VOLL,
                    p_min_pu=0,
                    carrier="shedding"
                )
