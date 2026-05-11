import pandas as pd


def get_embalses_closest_date(
    df: pd.DataFrame,
    target_date: str,
    date_col: str = "FECHA",
    water_cols: list[str] | None = None,
) -> pd.DataFrame:
    """
    Limpia la fecha del DataFrame de embalses y devuelve solo las filas
    correspondientes a la fecha disponible más cercana a target_date.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame original de embalses.

    target_date : str
        Fecha objetivo. Por ejemplo: "2014-01-01".

    date_col : str
        Nombre de la columna de fecha.

    water_cols : list[str] | None
        Columnas numéricas con coma decimal. Por defecto:
        ["AGUA_TOTAL", "AGUA_ACTUAL"].

    Returns
    -------
    pd.DataFrame
        DataFrame filtrado para la fecha más cercana disponible.
    """

    df = df.copy()

    if water_cols is None:
        water_cols = ["AGUA_TOTAL", "AGUA_ACTUAL"]

    # Limpiar fecha tipo:
    # Tue May 10 1988 02:00:00 GMT+0200 (Central European Summer Time)
    s = df[date_col].astype(str)

    # Eliminar la parte entre paréntesis
    s = s.str.replace(r"\s*\(.*\)$", "", regex=True)

    # Convertir a datetime
    df[date_col] = pd.to_datetime(
        s,
        format="%a %b %d %Y %H:%M:%S GMT%z",
        errors="coerce",
        utc=True
    )

    # Convertir a hora local de Madrid y quitar zona horaria
    df[date_col] = (
        df[date_col]
        .dt.tz_convert("Europe/Madrid")
        .dt.tz_localize(None)
    )

    # Quedarse solo con el día, sin hora
    df[date_col] = df[date_col].dt.normalize()

    # Convertir columnas de agua a float
    for col in water_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", ".", regex=False)
                .astype(float)
            )

    # Fecha objetivo
    target_date = pd.to_datetime(target_date).normalize()

    # Fechas disponibles
    available_dates = (
        df[date_col]
        .dropna()
        .drop_duplicates()
        .sort_values()
    )

    if available_dates.empty:
        raise ValueError("No hay fechas válidas en el DataFrame.")

    # Buscar fecha más cercana
    closest_date = available_dates.iloc[
        (available_dates - target_date).abs().argsort().iloc[0]
    ]

    print(f"Fecha objetivo: {target_date.date()}")
    print(f"Fecha disponible más cercana: {closest_date.date()}")

    # Filtrar DataFrame
    df_day = df[df[date_col] == closest_date].copy()

    return df_day




