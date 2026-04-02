import pypsa
import pandas as pd

def build_network(df_SYS_settings: pd.DataFrame) -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")
    start_date = str(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])
    horizon = str(df_SYS_settings.loc[3, "SYSTEM PARAMETERS"])
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24
    if horizon == "Static":
        grid.set_snapshots(pd.DatetimeIndex(["2026-01-01 00:00"]))

    elif horizon == "Multiperiod":
        grid.set_snapshots(pd.date_range(start_date, periods=simulation_hours, freq="h"))
  
    return grid
