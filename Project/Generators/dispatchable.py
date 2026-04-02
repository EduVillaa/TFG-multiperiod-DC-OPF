import pypsa
import pandas as pd

def add_dispatchable_generators(grid: pypsa.Network, df_Gen_Dispatchable: pd.DataFrame) -> None:

    df_Gen_Dispatchable["Pmin (MW)"] = pd.to_numeric(df_Gen_Dispatchable["Pmin (MW)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["a (€/MW²h)"] = pd.to_numeric(df_Gen_Dispatchable["a (€/MW²h)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["b (€/MWh)"] = pd.to_numeric(df_Gen_Dispatchable["b (€/MWh)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["c (€)"] = pd.to_numeric(df_Gen_Dispatchable["c (€)"], errors="coerce").fillna(0)
    df_Gen_Dispatchable["pwl segments"] = pd.to_numeric(df_Gen_Dispatchable["pwl segments"], errors="coerce").fillna(1).astype(int)

    for n in range(df_Gen_Dispatchable["GENERATOR LOCATION"].count()):
        Pmax = float(df_Gen_Dispatchable.loc[n, "Rated active power (MW)"])
        if pd.isna(Pmax):
            continue
        location = int(df_Gen_Dispatchable.loc[n, "GENERATOR LOCATION"])
        Pmin = float(df_Gen_Dispatchable.loc[n, "Pmin (MW)"])
        segs = int(df_Gen_Dispatchable.loc[n, "pwl segments"]) if pd.notna(df_Gen_Dispatchable.loc[n, "pwl segments"]) else 1

        a = df_Gen_Dispatchable.loc[n, "a (€/MW²h)"]
        b = df_Gen_Dispatchable.loc[n, "b (€/MWh)"]

        if segs > 1:
            step = Pmax / segs
            remaining_min = Pmin

            for i in range(segs):
                block_min_mw = max(0.0, min(step, remaining_min))
                remaining_min -= block_min_mw

                p_min_pu = block_min_mw / step  # p.u. del bloque

                P_mid = (i + 0.5) * step
                marginal_cost = 2 * a * P_mid + b

                grid.add(
                    "Generator", f"DispatchGen{location}_{n}_seg{i+1}",
                    bus=f"Bus_node_{location}",
                    p_nom=step,
                    p_min_pu=p_min_pu,
                    marginal_cost=marginal_cost,
                    carrier="AC",
                )
        else:
            grid.add(
                "Generator", f"DispatchGen{location}_{n}_seg1", #{n} es un indicador necesario para diferenciar los generadores que están en el mismo bus
                bus=f"Bus_node_{location}",
                p_nom=Pmax,
                p_min_pu=(Pmin / Pmax) if Pmax > 0 else 0.0,
                marginal_cost=b,
                carrier="AC",
            )
