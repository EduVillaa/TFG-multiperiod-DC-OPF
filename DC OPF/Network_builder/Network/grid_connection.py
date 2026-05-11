import pandas as pd
import pypsa
from pathlib import Path

def grid_connection(
    grid: pypsa.Network,
    df_Grid_connection: pd.DataFrame,
    df_TS_Energy_Prices: pd.DataFrame,
    df_SYS_settings: pd.DataFrame
) -> None:

    params = df_SYS_settings["SYSTEM PARAMETERS"]

    # Limpiar nombres de columnas por si vienen de Excel con espacios
    df_Grid_connection.columns = df_Grid_connection.columns.str.strip()

    # Convertir límites térmicos
    df_Grid_connection["Thermal limit (MW)"] = pd.to_numeric(
        df_Grid_connection["Thermal limit (MW)"],
        errors="coerce"
    ).fillna(1e6)



    # ============================================================
    # 1. AÑADIMOS LOS PCCs
    # ============================================================

    df_pcc = df_Grid_connection.dropna(
        subset=["Grid rated voltage at the PCC (kV)", "Longitude", "Latitude", "PCC name"]
    )

    for _, row in df_pcc.iterrows():
        busname = str(row["PCC name"]).strip()

        grid.add(
            "Bus",
            f"PCC_{busname}",
            v_nom=row["Grid rated voltage at the PCC (kV)"],
            x=row["Longitude"],
            y=row["Latitude"],
            carrier="AC"
        )

    # ============================================================
    # 2. AÑADIMOS LINKS ENTRE PCCs Y BUSES ESPAÑOLES
    # ============================================================

    df_links = df_Grid_connection.dropna(
        subset=["Bus", "PCC"]
    )

    for _, row in df_links.iterrows():
        pcc_name = str(row["PCC"]).strip()
        bus_name = str(row["Bus"]).strip()

        grid.add(
            "Link",
            f"LPCC_{pcc_name}_{bus_name}",
            bus0=f"PCC_{pcc_name}",
            bus1=f"Bus.{bus_name}",
            p_nom=row["Thermal limit (MW)"],
            p_min_pu=-1,
            p_max_pu=1,
            efficiency=1.0,
            marginal_cost=0,
            carrier="Interconnection"
        )

    # ============================================================
    # 3. AÑADIMOS GENERADORES DE IMPORTACIÓN / EXPORTACIÓN
    # ============================================================

    for _, row in df_pcc.iterrows():
        busname = str(row["PCC name"]).strip()

        external_market_capacity = df_links.loc[
            df_links["PCC"].astype(str).str.strip() == busname,
            "Thermal limit (MW)"
        ].sum()

        grid.add(
            "Generator",
            f"Grid_import_{busname}",
            bus=f"PCC_{busname}",
            p_nom=external_market_capacity,
            marginal_cost=0,
            carrier=f"Import_{busname}"
        )

        grid.add(
            "Generator",
            f"Grid_export_{busname}",
            bus=f"PCC_{busname}",
            p_nom=external_market_capacity,
            marginal_cost=-0,
            sign=-1,
            carrier=f"Export_{busname}"
        )

    link_name = "LPCC_Morocco_ES0 23"

    if link_name in grid.links.index:
        saldo_marruecos = morocco_net_exchange(grid, df_SYS_settings)
        saldo_marruecos = saldo_marruecos.reindex(grid.snapshots)

        p_nom_link = grid.links.loc[link_name, "p_nom"]

        saldo_marruecos = saldo_marruecos.clip(
            lower=-p_nom_link,
            upper=p_nom_link
        )

    grid.links_t.p_min_pu[link_name] = saldo_marruecos / p_nom_link
    grid.links_t.p_max_pu[link_name] = saldo_marruecos / p_nom_link

    # ============================================================
    # 4. PRECIOS HORARIOS SI ES MULTIPERIODO
    # ============================================================

    horizon = params["Static / Multiperiod"]

    if horizon != "Static":

        # PARA FRANCIA
        price_profile = read_prices_FR(df_SYS_settings, df_TS_Energy_Prices)

        for _, row in df_pcc.iterrows():
            busname = str(row["PCC name"]).strip()

            if busname == "France":
                grid.generators_t.marginal_cost[f"Grid_import_{busname}"] = price_profile
                grid.generators_t.marginal_cost[f"Grid_export_{busname}"] = -price_profile


def morocco_net_exchange(grid: pypsa.Network, df_SYS_settings: pd.DataFrame)-> pd.Series:
    params = df_SYS_settings["SYSTEM PARAMETERS"]

    start_Date = params["Start date (dd/mm/aaaa)"]
    horizon = params["Static / Multiperiod"]
    simulation_days = params["Simulation duration (days)"]

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        lista_df = []
        BASE_DIR = Path(__file__).resolve().parent.parent.parent

        for n in range(1, 12):
            nombre_doc = f"saldo_marruecos{n}.xls"
            ruta_doc = BASE_DIR / "System_data" / "Saldo_Marruecos_ESIOS" / nombre_doc
            tablas = pd.read_html(ruta_doc, header=0, decimal=",", thousands=".")
            df = tablas[0]
            df = df[["name", "value", "datetime"]]
            df = df[df["name"]=="Generación programada P48 Saldo Marruecos"]
            df = df[["datetime", "value"]]
            df["datetime"] = pd.to_datetime(df["datetime"], utc=True)
            df["datetime"] = df["datetime"].dt.tz_convert("Europe/Madrid").dt.tz_localize(None)
            df = df.set_index("datetime")
            df["value"] = pd.to_numeric(df["value"], errors="coerce")
            lista_df.append(df)

        df_final = pd.concat(lista_df)
        df_final = df_final.sort_index()
        df_final = df_final[~df_final.index.duplicated(keep="first")]
        start_Date = pd.to_datetime(start_Date)
        end = start_Date + pd.Timedelta(days=simulation_days)
        df_final = df_final.loc[(df_final.index >= start_Date) & (df_final.index < end)]

        saldo = df_final["value"]

        saldo = saldo.reindex(grid.snapshots)

        missing = saldo.isna().sum()
        total = len(saldo)

        print(f"Datos faltantes Marruecos después de reindexar: {missing}/{total} ({missing/total:.1%})")

        saldo = saldo.interpolate(method="time", limit=3)

        missing = saldo.isna().sum()
        total = len(saldo)

        print(f"Datos faltantes Marruecos después de interpolar: {missing}/{total} ({missing/total:.1%})")

        saldo = saldo.fillna(0)

        return saldo


def read_prices_FR(
    df_SYS_settings: pd.DataFrame,
    df_TS_Energy_Prices: pd.DataFrame
) -> pd.Series:

    params = df_SYS_settings["SYSTEM PARAMETERS"]

    start_Date = params["Start date (dd/mm/aaaa)"]
    horizon = params["Static / Multiperiod"]
    simulation_days = params["Simulation duration (days)"]
    simulation_hours = simulation_days * 24

    if horizon == "Static":
        return pd.Series([1])

    elif horizon == "Multiperiod":
        df = df_TS_Energy_Prices.loc[start_Date:].iloc[:simulation_hours]
        return df["Precio Francia (€/MWh)"]