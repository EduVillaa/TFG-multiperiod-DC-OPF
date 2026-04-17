import pypsa
import pandas as pd

def add_lines(grid: pypsa.Network, df_Net_Lines: pd.DataFrame) -> None:
    df_Net_Lines["Thermal limit (MW)"] = pd.to_numeric(
        df_Net_Lines["Thermal limit (MW)"], # Si se deja vacía la columna de límite térmico se asume que no hay límite.
        errors="coerce").fillna(1e6) 
    
    for n in range(df_Net_Lines["From"].count()):
        desde = int(df_Net_Lines.loc[n, "From"])
        hasta = int(df_Net_Lines.loc[n, "To"])
        grid.add(
            "Line", f"L{desde}{hasta}",
            bus0=f"Bus_node_{desde}",
            bus1=f"Bus_node_{hasta}",
            x=df_Net_Lines.loc[n, "Reactance (p.u)"],
            r=1e-6, #Para evitar el warning que sale al no incluir la resistencia
            s_nom=df_Net_Lines.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )
