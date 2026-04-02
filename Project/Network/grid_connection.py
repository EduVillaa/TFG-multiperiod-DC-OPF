import pandas as pd
import pypsa

def grid_connection(grid: pypsa.Network, df_Grid_connection: pd.DataFrame, df_TS_Energy_Prices: pd.DataFrame, df_SYS_settings: pd.DataFrame) -> None:
     
    grid.add(
        "Bus",
        "PCC",
        v_nom=df_Grid_connection.loc[0, "Grid rated voltage at the PCC"],
        carrier="AC"
    )

    df_Grid_connection["Thermal limit (MW)"] = pd.to_numeric(
        df_Grid_connection["Thermal limit (MW)"], errors="coerce"
    ).fillna(1e6) # Si se deja vacía la columna de límite térmico se asume que no hay límite.

    for n in range(df_Grid_connection["Bus"].count()):
        grid.add(
            "Line",
            f"LPCC{int(df_Grid_connection.loc[n, 'Bus'])}",
            bus0="PCC",
            bus1=f"Bus_node_{int(df_Grid_connection.loc[n, 'Bus'])}",
            x=df_Grid_connection.loc[n, "Reactance (p.u)"],
            r=1e-6,
            s_nom=df_Grid_connection.loc[n, "Thermal limit (MW)"],
            carrier="AC"
        )

    # Compra a red
    grid.add(
        "Generator",
        "Grid_import",
        bus="PCC",
        p_nom = df_Grid_connection.loc[0, "Import"],
        marginal_cost = df_Grid_connection.loc[1, "Import"], #Para OPF estático
        carrier="AC"
    )
    # Venta a red
    grid.add(
        "Generator",
        "Grid_export",
        bus="PCC",
        p_nom=df_Grid_connection.loc[0, "Export"],
        marginal_cost = -df_Grid_connection.loc[1, "Export"]*df_Grid_connection.loc[1, "Import"], #Para OPF estático
        #El coste marginal de exportación se da como porcentaje del coste marginal de importación
        sign=-1,
        carrier="AC"
    )

    #Los perfiles de tiempo temporales reescriben los estáticos si se realiza un OPF multiperiodo
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    if horizon != "Static":
        price_profile = read_prices(df_SYS_settings, df_TS_Energy_Prices)
        grid.generators_t.marginal_cost["Grid_import"] = price_profile
        grid.generators_t.marginal_cost["Grid_export"] = -price_profile*df_Grid_connection.loc[1, "Export"] #El coste marginal de exportación 
        #se da como porcentaje del coste marginal de importación


def read_prices(df_SYS_settings: pd.DataFrame,
                df_TS_Energy_Prices: pd.DataFrame) -> pd.Series:

    start_Date = pd.to_datetime(df_SYS_settings.loc[5, "SYSTEM PARAMETERS"])
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    simulation_days = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]
    simulation_hours = simulation_days*24

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        df = df_TS_Energy_Prices.loc[start_Date:].iloc[:simulation_hours]
        return df["Precio mercado SPOT Diario España (€/MWh)"]
