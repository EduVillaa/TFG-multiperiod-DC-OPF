import pypsa
import pandas as pd

def add_buses(grid: pypsa.Network, df_Net_Buses: pd.DataFrame) -> None:
    n_buses = df_Net_Buses["Bus rated voltage (kV)"].count()
    for n in range(n_buses):
        grid.add("Bus", f"Bus_node_{n+1}", v_nom=df_Net_Buses.loc[n, "Bus rated voltage (kV)"], carrier="AC")
