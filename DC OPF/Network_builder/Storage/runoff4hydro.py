import pandas as pd
from pathlib import Path
import xarray as xr
import numpy as np
import calendar


BASE_DIR = Path(__file__).resolve().parent


def get_node_runoff(
    lat: float,
    lon: float,
    df_runoff: pd.DataFrame,
    radius: float = 0.3,
    time_col: str = "valid_time",
    lat_col: str = "latitude",
    lon_col: str = "longitude",
    runoff_col: str = "ro",
) -> pd.DataFrame:
    """
    Devuelve la serie temporal de runoff asociada a un nodo usando la media
    local de ERA5-Land alrededor de sus coordenadas.

    La función selecciona todos los puntos de ERA5-Land dentro de una ventana:
        lat ± radius
        lon ± radius

    y calcula la media del runoff para cada instante temporal.

    Parameters
    ----------
    lat : float
        Latitud del nodo.
    lon : float
        Longitud del nodo.
    df_runoff : pd.DataFrame
        DataFrame con datos de runoff. Debe contener columnas de tiempo,
        latitud, longitud y runoff.
    radius : float
        Radio de la ventana espacial en grados.
        Por defecto 0.3.
    time_col : str
        Nombre de la columna temporal.
    lat_col : str
        Nombre de la columna de latitud.
    lon_col : str
        Nombre de la columna de longitud.
    runoff_col : str
        Nombre de la columna de runoff.

    Returns
    -------
    pd.DataFrame
        DataFrame con columnas:
        - valid_time
        - runoff
    """

    df = df_runoff.copy()
    df[time_col] = pd.to_datetime(df[time_col])

    # Filtrar puntos dentro de la ventana espacial alrededor del nodo
    mask = (
        (df[lat_col] >= lat - radius) &
        (df[lat_col] <= lat + radius) &
        (df[lon_col] >= lon - radius) &
        (df[lon_col] <= lon + radius)
    )

    df_local = df.loc[mask, [time_col, lat_col, lon_col, runoff_col]].copy()

    # Si no hay datos en la ventana, usar el punto más cercano como respaldo
    if df_local.empty or df_local[runoff_col].dropna().empty:
        coords = df[[lat_col, lon_col]].drop_duplicates().copy()

        coords["distance"] = np.sqrt(
            (coords[lat_col] - lat) ** 2 +
            (coords[lon_col] - lon) ** 2
        )

        nearest = coords.loc[coords["distance"].idxmin()]
        nearest_lat = nearest[lat_col]
        nearest_lon = nearest[lon_col]

        df_local = df[
            (df[lat_col] == nearest_lat) &
            (df[lon_col] == nearest_lon)
        ][[time_col, lat_col, lon_col, runoff_col]].copy()

        print(
            f"Aviso: no había datos válidos en la ventana local. "
            f"Se usa el punto más cercano: lat={nearest_lat}, lon={nearest_lon}"
        )

    # Media espacial de runoff para cada fecha
    result = (
        df_local
        .groupby(time_col, as_index=False)[runoff_col]
        .mean()
        .sort_values(time_col)
        .rename(columns={runoff_col: "runoff"})
        .reset_index(drop=True)
    )

    return result


def build_nodes_runoff_dataframe(
    node_coordinates: pd.DataFrame,
    df_runoff: pd.DataFrame,
    radius: float = 0.3,
    time_col: str = "valid_time",
    lat_col_nodes: str = "Latitude",
    lon_col_nodes: str = "Longitude",
    bus_col: str = "Bus name",
    lat_col_runoff: str = "latitude",
    lon_col_runoff: str = "longitude",
    runoff_col: str = "ro",
) -> pd.DataFrame:
    """
    Construye un DataFrame con la serie temporal de runoff para todos los nodos.

    Usa la función get_node_runoff() para calcular el runoff asociado a cada nodo.

    Parameters
    ----------
    node_coordinates : pd.DataFrame
        DataFrame con las coordenadas de los nodos.
        Debe contener columnas de longitud, latitud y nombre del bus.

    df_runoff : pd.DataFrame
        DataFrame de ERA5-Land con datos de runoff.
        Debe contener columnas de tiempo, latitud, longitud y runoff.

    radius : float
        Radio espacial usado por get_node_runoff(), en grados.

    time_col : str
        Nombre de la columna temporal.

    lat_col_nodes : str
        Nombre de la columna de latitud en node_coordinates.

    lon_col_nodes : str
        Nombre de la columna de longitud en node_coordinates.

    bus_col : str
        Nombre de la columna con el nombre del nodo/bus.

    lat_col_runoff : str
        Nombre de la columna de latitud en df_runoff.

    lon_col_runoff : str
        Nombre de la columna de longitud en df_runoff.

    runoff_col : str
        Nombre de la columna de runoff en df_runoff.

    Returns
    -------
    pd.DataFrame
        DataFrame con índice temporal y una columna por nodo.
    """

    result = None

    for _, row in node_coordinates.iterrows():
        bus = row[bus_col]
        lat = row[lat_col_nodes]
        lon = row[lon_col_nodes]

        print(f"Procesando nodo {bus}...")

        df_node = get_node_runoff(
            lat=lat,
            lon=lon,
            df_runoff=df_runoff,
            radius=radius,
            time_col=time_col,
            lat_col=lat_col_runoff,
            lon_col=lon_col_runoff,
            runoff_col=runoff_col,
        )

        df_node = df_node[[time_col, "runoff"]].copy()
        df_node = df_node.rename(columns={"runoff": bus})

        if result is None:
            result = df_node
        else:
            result = result.merge(df_node, on=time_col, how="outer")

    result[time_col] = pd.to_datetime(result[time_col])
    result = result.sort_values(time_col).reset_index(drop=True)

    return result


def build_runoff_factor_dataframe(
    df_runoff_nodes: pd.DataFrame,
    base_year: int = 2013,
    time_col: str = "valid_time",
    lower_limit: float | None = 0.3,
    upper_limit: float | None = 2.0,
) -> pd.DataFrame:
    """
    Convierte un DataFrame de runoff mensual por nodo en un DataFrame
    de factores de escala respecto al mismo mes del año base.

    Parameters
    ----------
    df_runoff_nodes : pd.DataFrame
        DataFrame con una columna temporal y una columna por nodo.
        Ejemplo:
            valid_time | ES0 0 | ES0 1 | PT0 1 | ...

    base_year : int
        Año respecto al cual se normalizan los factores.
        Por defecto, 2013.

    time_col : str
        Nombre de la columna temporal.

    lower_limit : float or None
        Límite inferior aplicado a los factores.
        Si es None, no se aplica límite inferior.

    upper_limit : float or None
        Límite superior aplicado a los factores.
        Si es None, no se aplica límite superior.

    Returns
    -------
    pd.DataFrame
        DataFrame con la misma estructura que df_runoff_nodes, pero con
        factores adimensionales.
    """

    df = df_runoff_nodes.copy()

    # Asegurar formato datetime
    df[time_col] = pd.to_datetime(df[time_col])

    # Columnas de nodos
    node_cols = [col for col in df.columns if col != time_col]

    # Crear DataFrame resultado
    result = df[[time_col]].copy()

    # Añadir año y mes auxiliares
    df["_year"] = df[time_col].dt.year
    df["_month"] = df[time_col].dt.month

    for node in node_cols:
        result[node] = np.nan

        for month in range(1, 13):
            # Valor base: mismo mes del año base
            base_mask = (df["_year"] == base_year) & (df["_month"] == month)
            base_values = df.loc[base_mask, node].dropna()

            if base_values.empty:
                print(f"Aviso: no hay valor base para {node}, mes {month}. Se usa factor 1.")
                base_value = np.nan
            else:
                base_value = base_values.iloc[0]

            # Filas del mismo mes en todos los años
            month_mask = df["_month"] == month

            if pd.isna(base_value) or base_value == 0:
                result.loc[month_mask, node] = 1.0
            else:
                result.loc[month_mask, node] = df.loc[month_mask, node] / base_value

    # Limpiar infinitos y NaN
    result[node_cols] = result[node_cols].replace([np.inf, -np.inf], np.nan)
    result[node_cols] = result[node_cols].fillna(1.0)

    # Limitar valores extremos si se desea
    if lower_limit is not None or upper_limit is not None:
        result[node_cols] = result[node_cols].clip(
            lower=lower_limit,
            upper=upper_limit
        )

    return result


def slice_runoff_factors(
    df_runoff_factors: pd.DataFrame,
    startdate: str,
    days: int,
    time_col: str = "valid_time",
) -> pd.DataFrame:
    """
    Recorta un DataFrame mensual de factores de runoff al periodo de simulación.

    Parameters
    ----------
    df_runoff_factors : pd.DataFrame
        DataFrame con una columna de fechas mensuales y una columna por nodo.

    startdate : str
        Fecha de inicio de simulación. Ejemplo: "2015-01-01".

    days : int
        Duración de la simulación en días.

    time_col : str
        Nombre de la columna temporal.

    Returns
    -------
    pd.DataFrame
        DataFrame recortado al periodo correspondiente.
    """

    df = df_runoff_factors.copy()

    df[time_col] = pd.to_datetime(df[time_col])

    start = pd.to_datetime(startdate)
    end = start + pd.Timedelta(days=days)

    # Como los factores son mensuales, necesitamos incluir todos los meses
    # que se solapan con el periodo de simulación.
    start_month = start.to_period("M").to_timestamp()
    end_month = (end - pd.Timedelta(seconds=1)).to_period("M").to_timestamp()

    mask = (
        (df[time_col] >= start_month) &
        (df[time_col] <= end_month)
    )

    result = df.loc[mask].copy()
    result = result.sort_values(time_col).reset_index(drop=True)

    return result


def scale_2013_hydro_inflow_with_monthly_weights(
    df_weights: pd.DataFrame,
    df_inflow_2013: pd.DataFrame,
    weights_time_col: str = "valid_time",
    inflow_time_col: str = "snapshot",
    hydro_suffix: str = " hydro",
) -> pd.DataFrame:
    """
    Multiplica los inflows horarios de 2013 por factores mensuales nodales.

    Parameters
    ----------
    df_weights : pd.DataFrame
        DataFrame mensual de factores de runoff.
        Estructura esperada:
            valid_time | ES0 0 | ES0 1 | ES0 2 | PT0 1 | ...

    df_inflow_2013 : pd.DataFrame
        DataFrame horario de inflows de 2013.
        Estructura esperada:
            snapshot | ES0 0 hydro | ES0 1 hydro | PT0 1 hydro | ...

    weights_time_col : str
        Nombre de la columna temporal del DataFrame de pesos.

    inflow_time_col : str
        Nombre de la columna temporal del DataFrame de inflow de 2013.

    hydro_suffix : str
        Sufijo usado en las columnas de inflow para obtener el nombre del nodo.
        Por defecto: " hydro".

    Returns
    -------
    pd.DataFrame
        DataFrame horario con índice temporal y columnas de unidades hydro.
        Las columnas son las mismas que en df_inflow_2013, excepto la columna temporal.
    """

    weights = df_weights.copy()
    inflow_2013 = df_inflow_2013.copy()

    # Convertir fechas
    weights[weights_time_col] = pd.to_datetime(weights[weights_time_col])
    inflow_2013[inflow_time_col] = pd.to_datetime(inflow_2013[inflow_time_col])

    # Poner índice temporal en inflow de 2013
    inflow_2013 = inflow_2013.set_index(inflow_time_col).sort_index()

    # Columnas de unidades hydro
    hydro_cols = inflow_2013.columns.tolist()

    # Fechas de inicio y final según el DataFrame de pesos
    start = weights[weights_time_col].min()

    last_month = weights[weights_time_col].max()
    last_year = last_month.year
    last_month_num = last_month.month
    last_day = calendar.monthrange(last_year, last_month_num)[1]

    end = pd.Timestamp(
        year=last_year,
        month=last_month_num,
        day=last_day,
        hour=23
    )

    # Índice horario objetivo
    target_snapshots = pd.date_range(
        start=start,
        end=end,
        freq="h"
    )

    # Crear diccionario de pesos: (year, month, node) -> factor
    weights = weights.set_index(weights_time_col).sort_index()

    # DataFrame resultado
    result = pd.DataFrame(
        index=target_snapshots,
        columns=hydro_cols,
        dtype=float
    )

    # Función auxiliar para obtener la hora equivalente de 2013
    def map_to_2013_timestamp(ts):
        """
        Convierte una fecha cualquiera a la misma fecha/hora en 2013.
        Si aparece 29 de febrero, usa 28 de febrero.
        """
        try:
            return pd.Timestamp(
                year=2013,
                month=ts.month,
                day=ts.day,
                hour=ts.hour,
                minute=ts.minute,
                second=ts.second
            )
        except ValueError:
            # Para 29 de febrero en años bisiestos
            return pd.Timestamp(
                year=2013,
                month=2,
                day=28,
                hour=ts.hour,
                minute=ts.minute,
                second=ts.second
            )

    # Índice equivalente en 2013 para cada snapshot objetivo
    source_snapshots_2013 = pd.DatetimeIndex([
        map_to_2013_timestamp(ts)
        for ts in target_snapshots
    ])

    for hydro_col in hydro_cols:
        # Obtener nodo a partir del nombre de la unidad
        # Ejemplo: "ES0 0 hydro" -> "ES0 0"
        if hydro_col.endswith(hydro_suffix):
            node = hydro_col.removesuffix(hydro_suffix)
        else:
            # Si no tiene el sufijo esperado, intenta usar la columna completa
            node = hydro_col

        if node not in weights.columns:
            print(f"Aviso: no hay pesos para el nodo '{node}'. Se usa factor 1.")
            monthly_factors = pd.Series(1.0, index=target_snapshots)
        else:
            # Factor mensual correspondiente a cada hora
            factor_values = []

            for ts in target_snapshots:
                month_timestamp = pd.Timestamp(
                    year=ts.year,
                    month=ts.month,
                    day=1
                )

                if month_timestamp in weights.index:
                    factor = weights.loc[month_timestamp, node]
                else:
                    factor = 1.0

                factor_values.append(factor)

            monthly_factors = pd.Series(
                factor_values,
                index=target_snapshots
            )

        # Inflow base de 2013 repetido sobre el periodo objetivo
        base_values = inflow_2013.loc[source_snapshots_2013, hydro_col].values

        # Multiplicar por factores mensuales
        result[hydro_col] = base_values * monthly_factors.values

    result.index.name = inflow_time_col

    return result


def build_hydro_inflow(
    base_dir: str | Path,
    startdate: str,
    days: int,
    radius: float = 0.3,
    base_year: int = 2013,
    lower_limit: float = 0.3,
    upper_limit: float = 2.0,
) -> pd.DataFrame:
    """
    Función principal.

    Lee los datos necesarios, calcula los factores de runoff y devuelve
    el inflow hidroeléctrico escalado para el periodo de simulación.

    Returns
    -------
    pd.DataFrame
        DataFrame horario con índice snapshot y columnas de unidades hydro.
    """

    base_dir = Path(base_dir)

    ruta_copernicus = base_dir / "System_data" / "climate_copernicus.nc"
    ruta_grid_inputs = base_dir / "GridInputs.xlsx"
    ruta_network_clean = base_dir / "System_data" / "network_clean.xlsx"

    # Leer ERA5-Land
    ds = xr.open_dataset(ruta_copernicus)
    df_runoff = ds["ro"].to_dataframe().reset_index()

    # Leer coordenadas de nodos
    node_coordinates = pd.read_excel(
        ruta_grid_inputs,
        sheet_name="Net_Buses",
        header=1,
        usecols="C:E"
    )

    # Leer inflows 2013
    df_2013_inflows = pd.read_excel(
        ruta_network_clean,
        sheet_name="storage_inflow"
    )

    # Construir runoff por nodo
    df_runoff_nodes = build_nodes_runoff_dataframe(
        node_coordinates=node_coordinates,
        df_runoff=df_runoff,
        radius=radius
    )

    # Calcular factores respecto a 2013
    df_runoff_factors = build_runoff_factor_dataframe(
        df_runoff_nodes=df_runoff_nodes,
        base_year=base_year,
        lower_limit=lower_limit,
        upper_limit=upper_limit
    )

    # Recortar al periodo simulado
    df_runoff_factors_sim = slice_runoff_factors(
        df_runoff_factors=df_runoff_factors,
        startdate=startdate,
        days=days
    )

    # Escalar inflows de 2013
    df_hydro_inflow_scaled = scale_2013_hydro_inflow_with_monthly_weights(
        df_weights=df_runoff_factors_sim,
        df_inflow_2013=df_2013_inflows,
        weights_time_col="valid_time",
        inflow_time_col="snapshot",
        hydro_suffix=" hydro"
    )

    return df_hydro_inflow_scaled


def get_hydro_inflow_node(
    df_hydro_inflow_scaled: pd.DataFrame,
    node_location: str,
) -> pd.Series:
    """
    Devuelve la serie de inflow de una unidad hydro concreta.
    """

    hydro_col = node_location + " hydro"

    if hydro_col not in df_hydro_inflow_scaled.columns:
        raise ValueError(f"No existe la columna {hydro_col}")

    return df_hydro_inflow_scaled[hydro_col]





