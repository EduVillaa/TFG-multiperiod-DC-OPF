import pypsa
import pandas as pd

def add_lines(grid: pypsa.Network, df_Net_Lines: pd.DataFrame) -> None:
    df_Net_Lines["Thermal limit (MW)"] = pd.to_numeric(
        df_Net_Lines["Thermal limit (MW)"], # Si se deja vacía la columna de límite térmico se asume que no hay límite.
        errors="coerce").fillna(9000) 
 
    for n in range(df_Net_Lines["From"].count()):
        desde_str = str(df_Net_Lines.loc[n, "From"])
        hasta_str = str(df_Net_Lines.loc[n, "To"])

        grid.add(
            "Line", f"L{desde_str}_{hasta_str}",
            bus0=f"Bus.{desde_str}",
            bus1=f"Bus.{hasta_str}",
            x=df_Net_Lines.loc[n, "Reactance (ohm)"],
            #r=1e-6, #Para evitar el warning que sale al no incluir la resistencia
            s_nom=df_Net_Lines.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )
    
