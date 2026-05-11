from Network_builder.Storage.runoff4hydro import (
    build_nodes_runoff_dataframe,
    build_runoff_factor_dataframe,
    slice_runoff_factors,
)

import pandas as pd
from pathlib import Path
import calendar
import xarray as xr


def scale_2013_ror_p_max_pu_with_monthly_weights(
    df_weights: pd.DataFrame,
    df_p_max_pu_2013: pd.DataFrame,
    weights_time_col: str = "valid_time",
    p_max_pu_time_col: str = "snapshot",
    ror_suffix: str = " ror",
) -> pd.DataFrame:
    """
    Multiplica los p_max_pu horarios de 2013 de los generadores ror
    por factores mensuales nodales de runoff.

    Parameters
    ----------
    df_weights : pd.DataFrame
        DataFrame mensual de factores de runoff.
        Estructura esperada:
            valid_time | ES0 0 | ES0 1 | ES0 2 | PT0 1 | ...

    df_p_max_pu_2013 : pd.DataFrame
        DataFrame horario de p_max_pu de 2013.
        Estructura esperada:
            snapshot | ES0 0 ror | ES0 1 ror | PT0 1 ror | ...

    weights_time_col : str
        Nombre de la columna temporal del DataFrame de pesos.

    p_max_pu_time_col : str
        Nombre de la columna temporal del DataFrame de p_max_pu de 2013.

    ror_suffix : str
        Sufijo usado en las columnas de p_max_pu para obtener el nombre del nodo.
        Por defecto: " ror".

    Returns
    -------
    pd.DataFrame
        DataFrame horario con índice temporal y columnas de generadores ror.
        Las columnas son las mismas que en df_p_max_pu_2013, excepto la columna temporal.
    """

    weights = df_weights.copy()
    p_max_pu_2013 = df_p_max_pu_2013.copy()

    # Convertir fechas
    weights[weights_time_col] = pd.to_datetime(weights[weights_time_col])
    p_max_pu_2013[p_max_pu_time_col] = pd.to_datetime(
        p_max_pu_2013[p_max_pu_time_col]
    )

    # Poner índice temporal en p_max_pu de 2013
    p_max_pu_2013 = p_max_pu_2013.set_index(p_max_pu_time_col).sort_index()

    # Columnas de generadores ror
    ror_cols = p_max_pu_2013.columns.tolist()

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
        hour=23,
    )

    # Índice horario objetivo
    target_snapshots = pd.date_range(
        start=start,
        end=end,
        freq="h",
    )

    # Crear índice mensual de factores
    weights = weights.set_index(weights_time_col).sort_index()

    # DataFrame resultado
    result = pd.DataFrame(
        index=target_snapshots,
        columns=ror_cols,
        dtype=float,
    )

    def map_to_2013_timestamp(ts: pd.Timestamp) -> pd.Timestamp:
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
                second=ts.second,
            )
        except ValueError:
            return pd.Timestamp(
                year=2013,
                month=2,
                day=28,
                hour=ts.hour,
                minute=ts.minute,
                second=ts.second,
            )

    # Índice equivalente en 2013 para cada snapshot objetivo
    source_snapshots_2013 = pd.DatetimeIndex(
        [map_to_2013_timestamp(ts) for ts in target_snapshots]
    )

    for ror_col in ror_cols:
        # Obtener nodo a partir del nombre del generador
        # Ejemplo: "ES0 0 ror" -> "ES0 0"
        if ror_col.endswith(ror_suffix):
            node = ror_col.removesuffix(ror_suffix)
        else:
            node = ror_col

        if node not in weights.columns:
            print(f"Aviso: no hay pesos para el nodo '{node}'. Se usa factor 1.")
            monthly_factors = pd.Series(1.0, index=target_snapshots)
        else:
            factor_values = []

            for ts in target_snapshots:
                month_timestamp = pd.Timestamp(
                    year=ts.year,
                    month=ts.month,
                    day=1,
                )

                if month_timestamp in weights.index:
                    factor = weights.loc[month_timestamp, node]
                else:
                    factor = 1.0

                factor_values.append(factor)

            monthly_factors = pd.Series(
                factor_values,
                index=target_snapshots,
            )

        # p_max_pu base de 2013 repetido sobre el periodo objetivo
        base_values = p_max_pu_2013.loc[source_snapshots_2013, ror_col].values

        # Multiplicar por factores mensuales y limitar a [0, 1]
        result[ror_col] = base_values * monthly_factors.values

    result = result.clip(lower=0, upper=1)

    result.index.name = p_max_pu_time_col

    return result


def build_ror_p_max_pu(
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
    el p_max_pu de los generadores ror escalado para el periodo de simulación.

    Returns
    -------
    pd.DataFrame
        DataFrame horario con índice snapshot y columnas de generadores ror.
    """

    base_dir = Path(base_dir)

    ruta_copernicus = base_dir / "System_data" / "climate_copernicus.nc"
    ruta_grid_inputs = base_dir / "GridInputs.xlsx"
    ruta_network_clean = base_dir / "System_data" / "network_clean.xlsx"

    # Leer ERA5-Land
    ds = xr.open_dataset(ruta_copernicus)
    df_runoff = ds["ro"].to_dataframe().reset_index()
    #print(df_runoff)
    # Leer coordenadas de nodos
    node_coordinates = pd.read_excel(
        ruta_grid_inputs,
        sheet_name="Net_Buses",
        header=1,
        usecols="C:E",
    )

    # Leer p_max_pu ror de 2013
    df_p_max_pu_2013 = pd.read_excel(
        ruta_network_clean,
        sheet_name="gen_p_max_pu",
        usecols="A,GX:HI",
    )
    #print(df_p_max_pu_2013)

    # Construir runoff por nodo
    df_runoff_nodes = build_nodes_runoff_dataframe(
        node_coordinates=node_coordinates,
        df_runoff=df_runoff,
        radius=radius,
    )
    #print(df_runoff_nodes)

    # Calcular factores respecto a 2013
    df_runoff_factors = build_runoff_factor_dataframe(
        df_runoff_nodes=df_runoff_nodes,
        base_year=base_year,
        lower_limit=lower_limit,
        upper_limit=upper_limit,
    )
    #print(df_runoff_factors)
    # Recortar al periodo simulado
    df_runoff_factors_sim = slice_runoff_factors(
        df_runoff_factors=df_runoff_factors,
        startdate=startdate,
        days=days,
    )

    # Escalar p_max_pu de 2013
    df_ror_p_max_pu_scaled = scale_2013_ror_p_max_pu_with_monthly_weights(
        df_weights=df_runoff_factors_sim,
        df_p_max_pu_2013=df_p_max_pu_2013,
        weights_time_col="valid_time",
        p_max_pu_time_col="snapshot",
        ror_suffix=" ror",
    )


    return df_ror_p_max_pu_scaled
