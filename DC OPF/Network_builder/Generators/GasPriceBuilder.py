import pandas as pd
from pathlib import Path

# ==========================================================================
# 1. OBTENEMOS SERIE TEMPORAL DIARIA DEL PRECIO DEL GAS EN ESPAÑA Y PORTUGAL
# ==========================================================================

def MIBGAS_prices(BASE_DIR):
    rutaMIBGAS_folder = BASE_DIR / "System_data" / "MIBGAS"

    df_list = []

    for n in range(2015, 2025):
        file_name = f"MIBGAS_Data_{n}.xlsx"
        ruta_file = rutaMIBGAS_folder / file_name

        if n==2023 or n==2024:
            df = pd.read_excel(ruta_file, sheet_name="MIBGAS Indexes", usecols="A,C,E")
            df.columns = ["time", "SPAIN GAS [EUR/MWh]", "PORTUGAL GAS [EUR/MWh]"]
            df["PORTUGAL GAS [EUR/MWh]"] = df["PORTUGAL GAS [EUR/MWh]"].fillna(df["SPAIN GAS [EUR/MWh]"])
            df["time"] = pd.to_datetime(df["time"])
            df = df.set_index("time")
            df_list.append(df)
            
        
        if n==2021 or n==2022:
            df = pd.read_excel(ruta_file, sheet_name="Indices", usecols="A,B,C")
            df.columns = ["time", "Area", "MIBGAS-ES-PT"]
            dfPT = df[df["Area"]=="PT"]
            dfPT = dfPT.rename(columns={"MIBGAS-ES-PT": "PORTUGAL GAS [EUR/MWh]"})
            dfPT["time"] = pd.to_datetime(dfPT["time"])
            dfPT = dfPT.set_index("time")

            dfSP = df[df["Area"]=="ES"]
            dfSP = dfSP.rename(columns={"MIBGAS-ES-PT": "SPAIN GAS [EUR/MWh]"})
            dfSP["time"] = pd.to_datetime(dfSP["time"])
            dfSP = dfSP.set_index("time")
            
            df_final = pd.concat([dfSP, dfPT], axis=1)
            df_final = df_final.drop("Area", axis=1)
            df_final["PORTUGAL GAS [EUR/MWh]"] = df_final["PORTUGAL GAS [EUR/MWh]"].fillna(df_final["SPAIN GAS [EUR/MWh]"])
            df_list.append(df_final)

        if n<2021:
            df = pd.read_excel(ruta_file, sheet_name="Indices", usecols= "A, C")
            df = df.rename(columns={df.columns[1]: "SPAIN GAS [EUR/MWh]"})
            df = df.rename(columns={"Delivery day": "time"})
            df["time"] = pd.to_datetime(df["time"])
            df = df.set_index("time")
            df["PORTUGAL GAS [EUR/MWh]"] = df["SPAIN GAS [EUR/MWh]"]
            df_list.append(df)

    df_gas_prices = pd.concat(df_list)      
    df_gas_prices = df_gas_prices.sort_index()
    return df_gas_prices

# ==========================================================================
# 2. OBTENEMOS SERIE TEMPORAL DE LOS COSTES DE EMISIÓN DE CO2 DE LA UE
# ==========================================================================

def carbon_emisions_cost(BASE_DIR):
    ruta_carbon_permits_prices = BASE_DIR / "System_data" / "EU_Carbon_Permits_Allowance.xlsx"

    df = pd.read_excel(ruta_carbon_permits_prices, usecols=["Date", "Price"])
    df = df.rename(columns={"Date": "time"})
    df["time"] = pd.to_datetime(df["time"])
    df = df.set_index("time")
    df = df.sort_index(ascending=True)
    return df

# ==========================================================================
# 3. TRATAMIENTO DE LOS DATAFRAMES DE LOS PASOS 2 Y 3
# ==========================================================================

def CCGT_dataframe_treatment(BASE_DIR, startdate, days):
    startdate = pd.to_datetime(startdate)
    end_date = startdate + pd.Timedelta(days=days)

    co2_price_full = carbon_emisions_cost(BASE_DIR)
    gas_price = MIBGAS_prices(BASE_DIR)

    co2_price_full.index = pd.to_datetime(co2_price_full.index)
    gas_price.index = pd.to_datetime(gas_price.index)

    co2_price_full = co2_price_full.sort_index()
    gas_price = gas_price.sort_index()

    # Si CO2 viene como DataFrame, nos quedamos con la primera columna
    if isinstance(co2_price_full, pd.DataFrame):
        co2_price_full = co2_price_full.iloc[:, 0]

    # Recortamos el gas al periodo de simulación
    gas_price = gas_price.loc[
        (gas_price.index >= startdate) &
        (gas_price.index < end_date)
    ].copy()

    if gas_price.empty:
        raise ValueError(
            f"No hay datos de gas entre {startdate} y {end_date}."
        )

    # MUY IMPORTANTE:
    # Reindexamos CO2 usando la serie completa, no la serie recortada
    co2_price = co2_price_full.reindex(
        gas_price.index,
        method="ffill"
    )

    # Si la simulación empieza antes del primer dato disponible de CO2,
    # usamos el primer dato posterior disponible
    co2_price = co2_price.bfill()

    if co2_price.isna().any():
        raise ValueError(
            "Siguen existiendo valores NaN en la serie de CO2. "
            "Comprueba que hay datos de CO2 cercanos al periodo simulado."
        )

    return gas_price, co2_price

# ==========================================================================
# 4. CALCULO DEL COSTE MARGINAL DEL CICLO COMBINADO
# ==========================================================================

def CCGT_marginal_cost(efficiency, gas_price, co2_price):
    EF_GAS = 0.202
    VOM_CCGT = 3.0

    ccgt_cost = pd.DataFrame(index=gas_price.index)
    ccgt_cost["SPAIN CCGT [EUR/MWh]"] = (
        gas_price["SPAIN GAS [EUR/MWh]"] / efficiency
        + EF_GAS * co2_price / efficiency
        + VOM_CCGT
    )

    ccgt_cost["PORTUGAL CCGT [EUR/MWh]"] = (
        gas_price["PORTUGAL GAS [EUR/MWh]"] / efficiency
        + EF_GAS * co2_price / efficiency
        + VOM_CCGT
    )

    return ccgt_cost

# ==========================================================================
# 5. CONVERTIMOS LA SERIE DIARIA DE PRECIO MARGINAL DEL CICLO COMBINADO EN SERIE HORARIA
# ==========================================================================

def daily_to_snapshots(daily_series: pd.Series, snapshots: pd.DatetimeIndex) -> pd.Series:
    daily_series = daily_series.copy()
    daily_series.index = pd.to_datetime(daily_series.index).normalize()
    daily_series = daily_series.sort_index()

    hourly = daily_series.reindex(snapshots.normalize(), method="ffill")
    hourly.index = snapshots

    return hourly



