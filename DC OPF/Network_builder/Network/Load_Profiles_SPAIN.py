import pandas as pd
from pathlib import Path
import re


def clean_region_name(x: str) -> str:
    """Convierte '01 Andalucía' en 'Andalucía'."""
    return re.sub(r"^\d+\s+", "", str(x)).strip()

def total_to_numeric(x):
    if pd.isna(x):
        return pd.NA

    x = str(x).strip()

    if x == ".":
        return pd.NA

    # Caso tipo "1.653.989" -> "1653.989"
    if x.count(".") > 1:
        parts = x.split(".")
        x = "".join(parts[:-1]) + "." + parts[-1]

    return pd.to_numeric(x, errors="coerce")

def convert_to_mwh(series: pd.Series, unit: str) -> pd.Series:
    """
    Convierte la serie de consumo anual a MWh.

    unit puede ser:
    - 'MWh'
    - 'GWh'
    - 'ktep'
    """

    unit = unit.lower()

    if unit == "mwh":
        return series

    if unit == "gwh":
        return series * 1000

    if unit == "ktep":
        return series * 11630  # 1 ktep ≈ 11.63 GWh = 11630 MWh

    raise ValueError("unit debe ser 'MWh', 'GWh' o 'ktep'.")


def read_demr_month(folder: str | Path, year: int, month: int) -> pd.DataFrame:
    """
    Lee un fichero DEMR_YYYYMM.
    """

    folder = Path(folder)
    ym = f"{year}{month:02d}"

    possible_paths = [
        folder / f"DEMR_{ym}",
        folder / f"DEMR_{ym}.csv",
        folder / f"DEMR_{ym}.txt",
    ]

    path = next((p for p in possible_paths if p.exists()), None)

    if path is None:
        raise FileNotFoundError(f"No se encontró el fichero DEMR_{ym}")

    df = pd.read_csv(
        path,
        sep=";",
        encoding="latin1",
        usecols=range(0, 6)
    )

    if "VERANO(1)/INVIERNO(0)" in df.columns:
        df = df.drop("VERANO(1)/INVIERNO(0)", axis=1)

    df["time"] = (
        pd.to_datetime(
            df["AÑO"].astype(str) + "-"
            + df["MES"].astype(str).str.zfill(2) + "-"
            + df["DIA"].astype(str).str.zfill(2)
        )
        + pd.to_timedelta(df["HORA"] - 1, unit="h")
    )

    df = df.rename(columns={"DEMANDA(MWh)": "demand_mwh"})

    df["demand_mwh"] = (
        df["demand_mwh"]
        .astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False)
    )

    df["demand_mwh"] = pd.to_numeric(df["demand_mwh"], errors="coerce")
    df = df.dropna(subset=["demand_mwh"])

    return df[["time", "demand_mwh"]]


def build_hourly_demand_by_region(
    df_consumos_anuales_CCAA: pd.DataFrame,
    demr_folder: str | Path,
    startdate: str,
    days: int,
    annual_unit: str = "ktep",
) -> pd.DataFrame:
    """
    Construye demanda horaria por comunidad autónoma.

    Parameters
    ----------
    df_consumos_anuales_CCAA:
        DataFrame con columnas:
        - 'Comunidades y Ciudades Autónomas'
        - 'Producto consumido'
        - 'Periodo'
        - 'Total'

    demr_folder:
        Carpeta donde están los ficheros DEMR_YYYYMM.

    startdate:
        Fecha inicial, por ejemplo '2022-01-01'.

    days:
        Duración de la simulación en días.

    annual_unit:
        Unidad del campo 'Total'. Probablemente 'ktep' para esos datos.
        Opciones: 'ktep', 'GWh', 'MWh'.

    Returns
    -------
    DataFrame:
        Columna 'time' + una columna por comunidad autónoma.
    """

    startdate = pd.to_datetime(startdate)
    enddate = startdate + pd.Timedelta(days=days)

    years_needed = range(startdate.year, enddate.year + 1)

    # ------------------------------------------------------------
    # 1. Preparar consumo anual por CCAA
    # ------------------------------------------------------------

    annual = df_consumos_anuales_CCAA.copy()

    annual = annual[annual["Producto consumido"] == "Electricidad"].copy()

    annual["region"] = annual["Comunidades y Ciudades Autónomas"].apply(clean_region_name)
    annual["year"] = annual["Periodo"].astype(int)
    annual["Total"] = annual["Total"].apply(total_to_numeric)

    annual = annual.dropna(subset=["Total"])

    annual["annual_mwh"] = convert_to_mwh(annual["Total"], annual_unit)

    annual_pivot = annual.pivot_table(
        index="year",
        columns="region",
        values="annual_mwh",
        aggfunc="sum"
    )

    # ------------------------------------------------------------
    # 2. Leer demanda horaria nacional de los años necesarios
    # ------------------------------------------------------------

    hourly_parts = []

    for year in years_needed:
        for month in range(1, 13):
            try:
                hourly_parts.append(read_demr_month(demr_folder, year, month))
            except FileNotFoundError:
                pass

    if not hourly_parts:
        raise ValueError("No se ha leído ningún fichero DEMR.")

    national = pd.concat(hourly_parts, ignore_index=True)
    national = national.sort_values("time").reset_index(drop=True)
    national["year"] = national["time"].dt.year

    # Nos quedamos con años completos para calcular bien el reparto anual
    national_year_total = national.groupby("year")["demand_mwh"].sum()
    # ------------------------------------------------------------
    # 3. Construir demanda horaria por comunidad usando pesos regionales
    # ------------------------------------------------------------

    result = national.copy()

    # Si tu demanda nacional es PENINSULAR, elimina islas, Ceuta y Melilla
    regions_to_exclude = [
        "Balears, Illes",
        "Canarias",
        "Ceuta",
        "Melilla"
    ]

    annual_pivot = annual_pivot.drop(
        columns=[c for c in regions_to_exclude if c in annual_pivot.columns],
        errors="ignore"
    )

    for region in annual_pivot.columns:
        result[region] = pd.NA

    for year in result["year"].unique():

        if year not in annual_pivot.index:
            continue

        regional_consumption = annual_pivot.loc[year].dropna()

        total_regional_consumption = regional_consumption.sum()

        if total_regional_consumption == 0:
            continue

        mask = result["year"] == year

        for region in regional_consumption.index:

            regional_weight = regional_consumption[region] / total_regional_consumption

            result.loc[mask, region] = (
                result.loc[mask, "demand_mwh"] * regional_weight
            )

    debug = result[
        (result["time"] >= "2022-04-30 20:00")
        & (result["time"] <= "2022-05-01 06:00")
    ].copy()

    region_cols = [
        c for c in result.columns
        if c not in ["time", "demand_mwh", "year"]
    ]

    debug["sum_regions"] = debug[region_cols].sum(axis=1)

    if "Madrid, Comunidad de" in debug.columns:
        debug["w_madrid"] = debug["Madrid, Comunidad de"] / debug["demand_mwh"]


    # ------------------------------------------------------------
    # 4. Filtrar horizonte solicitado
    # ------------------------------------------------------------

    result = result[
        (result["time"] >= startdate)
        & (result["time"] < enddate)
    ].copy()

    result = result.drop(columns=["demand_mwh", "year"])

    return result.reset_index(drop=True)


def build_monthly_nodal_load_weights_ES(
    df_pypsa_load_profiles: pd.DataFrame,
    node_to_region: dict,
    exclude_portugal: bool = True,
) -> pd.DataFrame:
    """
    Calcula pesos nodales mensuales de demanda a partir de perfiles horarios PyPSA 2013.

    Para cada comunidad c, mes m y nodo n:

        w_{n,m} = E_{n,m}^{2013} / E_{c,m}^{2013}

    donde:
        E_{n,m}^{2013} = suma mensual de demanda del nodo n
        E_{c,m}^{2013} = suma mensual de demanda de todos los nodos de la comunidad c

    Returns
    -------
    DataFrame con columnas:
        Region, Month, Node, Weight
    """

    df = df_pypsa_load_profiles.copy()

    # Detectar columna temporal
    if "snapshot" in df.columns:
        time_col = "snapshot"
    elif "time" in df.columns:
        time_col = "time"
    else:
        raise ValueError("No se encuentra columna temporal: debe llamarse 'snapshot' o 'time'.")

    df[time_col] = pd.to_datetime(df[time_col])
    df["Month"] = df[time_col].dt.month

    # Nodos disponibles en el DataFrame y en el diccionario
    node_cols = [c for c in df.columns if c not in [time_col, "Month"]]

    if exclude_portugal:
        node_cols = [n for n in node_cols if not str(n).startswith("PT")]

    missing_nodes = [n for n in node_cols if n not in node_to_region]

    if missing_nodes:
        raise ValueError(
            "Hay nodos en df_pypsa_load_profiles que no están en node_to_region: "
            f"{missing_nodes}"
        )

    # Pasar a formato largo
    long_df = df[[time_col, "Month"] + node_cols].melt(
        id_vars=[time_col, "Month"],
        var_name="Node",
        value_name="Demand"
    )

    long_df["Region"] = long_df["Node"].map(node_to_region)

    if exclude_portugal:
        long_df = long_df[~long_df["Region"].astype(str).str.contains(r"\(PT\)", regex=True)]

    long_df["Demand"] = pd.to_numeric(long_df["Demand"], errors="coerce").fillna(0)

    # Energía mensual por nodo
    node_month = (
        long_df
        .groupby(["Region", "Month", "Node"], as_index=False)["Demand"]
        .sum()
        .rename(columns={"Demand": "Node monthly demand"})
    )

    # Energía mensual total por comunidad
    region_month = (
        node_month
        .groupby(["Region", "Month"], as_index=False)["Node monthly demand"]
        .sum()
        .rename(columns={"Node monthly demand": "Region monthly demand"})
    )

    weights = node_month.merge(
        region_month,
        on=["Region", "Month"],
        how="left"
    )

    weights["Weight"] = (
        weights["Node monthly demand"] /
        weights["Region monthly demand"]
    )

    return weights[[
        "Region",
        "Month",
        "Node",
        "Node monthly demand",
        "Region monthly demand",
        "Weight"
    ]]




node_to_region = {
    "ES0 0": "Madrid",
    "ES0 1": "Cataluña",
    "ES0 10": "Valencia",
    "ES0 11": "Cantabria",
    "ES0 12": "Castilla y León",
    "ES0 13": "Andalucía",
    "ES0 14": "Castilla y León",
    "ES0 15": "Cataluña",
    "ES0 16": "Castilla y León",
    "ES0 17": "Andalucía",
    "ES0 18": "Andalucía",
    "ES0 19": "Asturias",
    "ES0 2": "Galicia",
    "ES0 20": "Castilla la Mancha",
    "ES0 21": "Murcia",
    "ES0 22": "Madrid",
    "ES0 23": "Andalucía",
    "ES0 24": "Navarra",
    "ES0 25": "Castilla y León",
    "ES0 26": "Andalucía",
    "ES0 27": "Cataluña",
    "ES0 28": "La Rioja",
    "ES0 29": "Valencia",
    "ES0 3": "País Vasco",
    "ES0 30": "Cataluña",
    "ES0 31": "Galicia",
    "ES0 32": "Cataluña",
    "ES0 33": "Extremadura",
    "ES0 34": "Extremadura",
    "ES0 35": "Extremadura",
    "ES0 36": "Galicia",
    "ES0 37": "País Vasco",
    "ES0 38": "Castilla la Mancha",
    "ES0 39": "Aragón",
    "ES0 4": "Valencia",
    "ES0 40": "Castilla y León",
    "ES0 5": "Andalucía",
    "ES0 6": "Aragón",
    "ES0 7": "Castilla y León",
    "ES0 8": "Andalucía",
    "ES0 9": "Galicia",
}


"""
check = (
    df_monthly_node_weights
    .groupby(["Region", "Month"])["Weight"]
    .sum()
    .reset_index()
)

print(check)"""



def build_hourly_nodal_demand(
    df_demanda_ccaa: pd.DataFrame,
    df_monthly_node_weights: pd.DataFrame,
) -> pd.DataFrame:

    demand = df_demanda_ccaa.copy()
    weights = df_monthly_node_weights.copy()

    rename_regions = {
        "Madrid, Comunidad de": "Madrid",
        "Murcia, Región de": "Murcia",
        "Navarra, Comunidad Foral de": "Navarra",
        "Rioja, La": "La Rioja",
        "Castilla - La Mancha": "Castilla la Mancha",
        "Comunitat Valenciana": "Valencia",
        "Asturias, Principado de": "Asturias",
    }

    demand = demand.rename(columns=rename_regions)

    demand["time"] = pd.to_datetime(demand["time"])
    demand["Month"] = demand["time"].dt.month

    region_cols = [c for c in demand.columns if c not in ["time", "Month"]]

    # asegurar numérico
    for c in region_cols:
        demand[c] = pd.to_numeric(demand[c], errors="coerce")

    demand_long = demand.melt(
        id_vars=["time", "Month"],
        value_vars=region_cols,
        var_name="Region",
        value_name="Region demand"
    )

    weights["Weight"] = pd.to_numeric(weights["Weight"], errors="coerce")

    """# DEBUG: regiones que no cruzan
    demand_regions = set(demand_long["Region"].unique())
    weight_regions = set(weights["Region"].unique())

    print("Regiones en demanda pero no en pesos:")
    print(sorted(demand_regions - weight_regions))

    print("Regiones en pesos pero no en demanda:")
    print(sorted(weight_regions - demand_regions))"""

    merged = demand_long.merge(
        weights[["Region", "Month", "Node", "Weight"]],
        on=["Region", "Month"],
        how="inner"
    )

    merged["Node demand"] = merged["Region demand"] * merged["Weight"]

    nodal_demand = merged.pivot_table(
        index="time",
        columns="Node",
        values="Node demand",
        aggfunc="sum"
    ).reset_index()

    nodal_demand.columns.name = None

    node_cols = [c for c in nodal_demand.columns if c != "time"]

    node_cols_sorted = sorted(
        node_cols,
        key=lambda x: int(str(x).split()[1]) if str(x).startswith("ES0 ") else str(x)
    )

    nodal_demand = nodal_demand[["time"] + node_cols_sorted]

    return nodal_demand


