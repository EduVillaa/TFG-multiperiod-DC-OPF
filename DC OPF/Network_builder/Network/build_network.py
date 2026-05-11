import pypsa
import pandas as pd

def build_network(df_SYS_settings: pd.DataFrame) -> pypsa.Network:
    grid = pypsa.Network()
    grid.add("Carrier", "AC")

    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    static_snapshot_datetime = params["Static snapshot datetime"]
    
    if horizon == "Static":
        grid.set_snapshots(pd.DatetimeIndex([static_snapshot_datetime]))

    elif horizon == "Multiperiod":
        start_date = params["Start date (dd/mm/aaaa)"]
        simulation_days = params["Simulation duration (days)"]
        simulation_hours = int(simulation_days) * 24

        grid.set_snapshots(
            pd.date_range(start_date, periods=simulation_hours, freq="h")
        )

    else:
        raise ValueError(f"Valor no válido para 'Static / Multiperiod': {horizon}")

    return grid