import pandas as pd


def renewable_profile_builder(
    df_installed_capacity: pd.DataFrame,
    df_node_weights: pd.DataFrame,
    df_profiles: pd.DataFrame,
    days: int,
    startdate: str,
) -> pd.DataFrame:
   

    # Copias para evitar modificar los originales
    profiles = df_profiles.copy()
    installed_capacity = df_installed_capacity.copy()
    node_weights = df_node_weights.copy()

    # Convertir peso global en peso relativo dentro de cada región
    node_weights["Rated active power (normalized)"] = pd.to_numeric(
        node_weights["Rated active power (normalized)"],
        errors="coerce"
    )

    regional_sum = node_weights.groupby("Region")["Rated active power (normalized)"].transform("sum")

    if (regional_sum == 0).any():
        bad_regions = node_weights.loc[regional_sum == 0, "Region"].unique()
        raise ValueError(f"Hay regiones con suma de pesos igual a cero: {bad_regions}")

    node_weights["Regional weight"] = (
        node_weights["Rated active power (normalized)"] / regional_sum
    )

    # Asegurar formato datetime
    profiles["time"] = pd.to_datetime(profiles["time"]).dt.round("h")
    startdate = pd.to_datetime(startdate)
    enddate = startdate + pd.Timedelta(days=days)

    # Filtrar horizonte temporal
    profiles = profiles[
        (profiles["time"] >= startdate) &
        (profiles["time"] < enddate)
    ].copy()

    if profiles.empty:
        raise ValueError("No hay datos de perfiles renovables para el periodo seleccionado.")

    # Preparar potencia instalada
    installed_capacity = installed_capacity.rename(columns={"Year/Region": "year"})
    installed_capacity["year"] = installed_capacity["year"].astype(int)
    installed_capacity = installed_capacity.set_index("year")

    # Crear dataframe de salida
    result = pd.DataFrame(index=profiles.index)
    result["time"] = profiles["time"].values

    # Año de cada snapshot
    years = profiles["time"].dt.year

    # Iterar por cada nodo renovable
    for _, row in node_weights.iterrows():

        node = row["GENERATOR LOCATION"]
        region = row["Region"]
        weight = row["Regional weight"]

        # Comprobaciones
        if pd.isna(region) or pd.isna(weight):
            continue

        if region not in profiles.columns:
            raise ValueError(f"La región '{region}' no existe en df_profiles.")

        if region not in installed_capacity.columns:
            raise ValueError(f"La región '{region}' no existe en df_installed_capacity.")

        # Potencia instalada correspondiente al año de cada snapshot
        capacity = years.map(installed_capacity[region])

        if capacity.isna().any():
            missing_years = sorted(years[capacity.isna()].unique())
            raise ValueError(
                f"No hay potencia instalada para la región '{region}' "
                f"en los años {missing_years}."
            )

        # Perfil nodal horario
        result[node] = profiles[region].values * capacity.values * weight

    return result.reset_index(drop=True)






