import pandas as pd

def regional_hourly_demand_builder(
    df_regional_demand: pd.DataFrame,
    df_total_hourly_demand: pd.DataFrame,
    startdate: str,
    days: int,
) -> pd.DataFrame:
    """
    Desagrega la demanda horaria total de Portugal por región.

    Parameters
    ----------
    df_regional_demand : pd.DataFrame
        Tabla anual con columnas:
        - "Year/Region"
        - regiones, por ejemplo: "Norte (PT)", "Centro (PT)", ...

        Los valores pueden estar en kWh, GWh, etc. Solo se usan para calcular pesos.

    df_total_hourly_demand : pd.DataFrame
        Tabla horaria con columnas:
        - "time"
        - "Actual Load (MW)"

    startdate : str
        Fecha inicial de simulación, por ejemplo "2022-01-01".

    days : int
        Número de días de simulación.

    Returns
    -------
    pd.DataFrame
        DataFrame con:
        - time
        - una columna por región con demanda horaria en MW
    """

    regional = df_regional_demand.copy()
    total = df_total_hourly_demand.copy()

    # Asegurar formato datetime
    total["time"] = pd.to_datetime(total["time"])

    start = pd.to_datetime(startdate)
    end = start + pd.Timedelta(days=days)

    # Filtrar periodo de simulación
    total = total[(total["time"] >= start) & (total["time"] < end)].copy()

    if total.empty:
        raise ValueError("No hay datos horarios dentro del periodo seleccionado.")

    # Detectar columnas regionales
    year_col = "Year/Region"
    load_col = "Actual Load (MW)"

    if year_col not in regional.columns:
        raise ValueError(f"df_regional_demand debe contener la columna '{year_col}'.")

    if load_col not in total.columns:
        raise ValueError(f"df_total_hourly_demand debe contener la columna '{load_col}'.")

    region_cols = [col for col in regional.columns if col != year_col]

    # Convertir año a entero
    regional[year_col] = regional[year_col].astype(int)

    # Calcular pesos anuales por región
    regional_weights = regional.copy()
    regional_weights[region_cols] = regional_weights[region_cols].astype(float)

    regional_weights["total_regional"] = regional_weights[region_cols].sum(axis=1)

    for col in region_cols:
        regional_weights[col] = regional_weights[col] / regional_weights["total_regional"]

    regional_weights = regional_weights.drop(columns="total_regional")

    # Añadir año a la demanda horaria
    total["year"] = total["time"].dt.year

    # Unir cada hora con los pesos de su año
    merged = total.merge(
        regional_weights,
        left_on="year",
        right_on=year_col,
        how="left"
    )

    if merged[region_cols].isna().any().any():
        missing_years = merged.loc[
            merged[region_cols].isna().any(axis=1), "year"
        ].unique()

        raise ValueError(
            f"No hay datos regionales para estos años: {missing_years}"
        )

    # Multiplicar demanda total por peso regional
    result = pd.DataFrame()
    result["time"] = merged["time"]

    for region in region_cols:
        result[region] = merged[load_col] * merged[region]

    return result


def build_monthly_nodal_load_weights_PT(
    df_pypsa_load_profiles: pd.DataFrame,
    node_to_region: dict,
    exclude_spain: bool = True,
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

    if exclude_spain:
        node_cols = [n for n in node_cols if not str(n).startswith("ES")]

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

    if exclude_spain:
        long_df = long_df[~long_df["Region"].astype(str).str.contains(r"\(ES\)", regex=True)]

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


node_to_region_PT = {
    "PT0 0": "Lisboa (PT)",
    "PT0 1": "Norte (PT)",
    "PT0 2": "Centro (PT)",
    "PT0 3": "Algarve (PT)",
    "PT0 4": "Centro (PT)",
    "PT0 5": "Alentejo (PT)",
    "PT0 6": "Centro (PT)",
    "PT0 7": "Norte (PT)",
}


def build_hourly_nodal_demand_PT(
    df_demand_PT_regions: pd.DataFrame,
    df_monthly_node_weights: pd.DataFrame,
) -> pd.DataFrame:

    demand = df_demand_PT_regions.copy()
    weights = df_monthly_node_weights.copy()


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




