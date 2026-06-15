import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def build_node_curtailment_dataframe(
    curtailment_df: pd.DataFrame,
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
) -> pd.DataFrame | None:
    """
    Construye un DataFrame con el curtailment agregado por nodo.

    A partir de columnas del tipo:
        PV_ES0 0_curtailment
        Wind_ES0 18_curtailment

    devuelve un DataFrame:
        index  -> snapshots
        columns -> nodos
        values -> curtailment PV + Wind por nodo

    Parameters
    ----------
    curtailment_df : pd.DataFrame
        DataFrame con columnas de curtailment.
    technologies : tuple[str, ...]
        Tecnologías que se quieren incluir.
    snapshot_col : str
        Nombre de la columna temporal, si existe.

    Returns
    -------
    node_curtailment : pd.DataFrame | None
        Curtailment agregado por nodo.
    """

    df = curtailment_df.copy()

    # Asegurar índice temporal
    if snapshot_col in df.columns:
        df[snapshot_col] = pd.to_datetime(df[snapshot_col])
        df = df.set_index(snapshot_col)
    else:
        df.index = pd.to_datetime(df.index)

    # Patrón: PV_ES0 0_curtailment o Wind_ES0 18_curtailment
    tech_pattern = "|".join(technologies)
    pattern = re.compile(rf"^({tech_pattern})_(.+)_curtailment$")

    node_series = {}

    for col in df.columns:
        match = pattern.match(col)

        if match is None:
            continue

        node = match.group(2)

        if node not in node_series:
            node_series[node] = df[col].astype(float)
        else:
            node_series[node] = node_series[node] + df[col].astype(float)

    if not node_series:
        return None

    node_curtailment = pd.DataFrame(node_series, index=df.index)

    # Limpieza
    node_curtailment = node_curtailment.replace([np.inf, -np.inf], np.nan)
    node_curtailment = node_curtailment.dropna(axis=1, how="all")
    node_curtailment = node_curtailment.fillna(0)

    # Eliminar nodos sin curtailment
    node_curtailment = node_curtailment.loc[:, node_curtailment.sum(axis=0) > 0]

    if node_curtailment.empty:
        return None

    return node_curtailment


def plot_node_curtailment_heatmap(
    curtailment_df: pd.DataFrame,
    horizon: str = "Multiperiod",
    resolution: str = "auto",
    sort_by: str = "sum",
    top_n_nodes: int | None = 15,
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
):
    """
    Devuelve un heatmap del curtailment PV + Wind por nodo.

    Parameters
    ----------
    curtailment_df : pd.DataFrame
        DataFrame con columnas del tipo:
        PV_ES0 0_curtailment, Wind_ES0 18_curtailment, etc.

    horizon : str
        Solo genera figura si horizon == "Multiperiod".

    resolution : str
        Resolución temporal del heatmap:
        - "hourly" : valores horarios
        - "daily"  : curtailment diario acumulado
        - "weekly" : curtailment semanal acumulado
        - "auto"   : selecciona resolución según duración

    sort_by : str
        Criterio para ordenar los nodos:
        - "sum"  : mayor curtailment total
        - "max"  : mayor curtailment máximo
        - "mean" : mayor curtailment medio
        - "name" : orden alfabético

    top_n_nodes : int | None
        Número máximo de nodos a mostrar.
        Si es None, muestra todos.

    technologies : tuple[str, ...]
        Tecnologías incluidas en el cálculo.

    snapshot_col : str
        Nombre de la columna temporal, si existe.

    Returns
    -------
    fig : matplotlib.figure.Figure | None
        Figura generada o None si no hay datos.
    """

    if horizon != "Multiperiod":
        return None

    node_curtailment = build_node_curtailment_dataframe(
        curtailment_df=curtailment_df,
        technologies=technologies,
        snapshot_col=snapshot_col,
    )

    if node_curtailment is None or node_curtailment.empty:
        return None

    if len(node_curtailment.index) < 2:
        return None

    # Selección automática de resolución
    if resolution == "auto":
        duration_days = (
            node_curtailment.index.max() - node_curtailment.index.min()
        ).total_seconds() / 86400

        if duration_days <= 14:
            resolution = "hourly"
        elif duration_days <= 120:
            resolution = "daily"
        else:
            resolution = "weekly"

    # Agregación temporal
    if resolution == "hourly":
        plot_data = node_curtailment.copy()
        title_prefix = "Hourly"
        cbar_label = "Curtailment [MW]"
        dt_days = None

    elif resolution == "daily":
        # Al ser datos horarios en MW, la suma diaria equivale a MWh si cada snapshot es 1 h
        plot_data = node_curtailment.resample("D").sum()
        title_prefix = "Daily accumulated"
        cbar_label = "Curtailment [MWh/day]"
        dt_days = 1.0

    elif resolution == "weekly":
        plot_data = node_curtailment.resample("W").sum()
        title_prefix = "Weekly accumulated"
        cbar_label = "Curtailment [MWh/week]"
        dt_days = 7.0

    else:
        raise ValueError(
            "resolution debe ser 'hourly', 'daily', 'weekly' o 'auto'"
        )

    plot_data = plot_data.dropna(axis=1, how="all")
    plot_data = plot_data.loc[:, plot_data.sum(axis=0) > 0]

    if plot_data.empty:
        return None

    # Ordenar nodos
    if sort_by == "sum":
        ordered_cols = plot_data.sum(axis=0).sort_values(ascending=False).index
    elif sort_by == "max":
        ordered_cols = plot_data.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = plot_data.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(plot_data.columns)
    else:
        ordered_cols = plot_data.columns

    plot_data = plot_data[ordered_cols]

    # Quedarse con los top N nodos
    if top_n_nodes is not None:
        plot_data = plot_data.iloc[:, :top_n_nodes]

    if plot_data.empty:
        return None

    data = plot_data.T.values

    x_num = mdates.date2num(plot_data.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    if dt_days is None:
        dt_days = np.median(np.diff(x_num))

    x_min = x_num[0] - dt_days / 2
    x_max = x_num[-1] + dt_days / 2

    fig_height = max(4, 0.45 * len(plot_data.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(plot_data.columns) - 0.5],
        interpolation="nearest",
        vmin=0,
    )

    ax.set_yticks(np.arange(len(plot_data.columns)))
    ax.set_yticklabels(plot_data.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.set_title(
        f"{title_prefix} PV + wind curtailment heatmap - "
        f"Top {len(plot_data.columns)} nodes"
    )
    ax.set_xlabel("Time")
    ax.set_ylabel("Nodes")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label(cbar_label)

    fig.tight_layout()

    return fig


def plot_node_curtailment_heatmap_hourly(
    curtailment_df: pd.DataFrame,
    horizon: str = "Multiperiod",
    sort_by: str = "sum",
    top_n_nodes: int | None = 15,
):
    """
    Heatmap horario del curtailment PV + Wind por nodo.
    """

    return plot_node_curtailment_heatmap(
        curtailment_df=curtailment_df,
        horizon=horizon,
        resolution="hourly",
        sort_by=sort_by,
        top_n_nodes=top_n_nodes,
    )


def plot_node_curtailment_heatmap_daily(
    curtailment_df: pd.DataFrame,
    horizon: str = "Multiperiod",
    sort_by: str = "sum",
    top_n_nodes: int | None = 15,
):
    """
    Heatmap del curtailment diario acumulado PV + Wind por nodo.
    """

    return plot_node_curtailment_heatmap(
        curtailment_df=curtailment_df,
        horizon=horizon,
        resolution="daily",
        sort_by=sort_by,
        top_n_nodes=top_n_nodes,
    )


def plot_node_curtailment_heatmap_weekly(
    curtailment_df: pd.DataFrame,
    horizon: str = "Multiperiod",
    sort_by: str = "sum",
    top_n_nodes: int | None = 15,
):
    """
    Heatmap del curtailment semanal acumulado PV + Wind por nodo.
    """

    return plot_node_curtailment_heatmap(
        curtailment_df=curtailment_df,
        horizon=horizon,
        resolution="weekly",
        sort_by=sort_by,
        top_n_nodes=top_n_nodes,
    )




def curtailment_heatmap_resolution_choice(
    df_SYS_settings: pd.DataFrame,
    curtailment_df: pd.DataFrame,
    top_n_nodes: int | None = 15,
):
    """
    Selecciona la resolución del heatmap de curtailment PV + Wind
    en función de los parámetros de simulación.

    Parameters
    ----------
    df_SYS_settings : pd.DataFrame
        DataFrame con la columna 'SYSTEM PARAMETERS'.

    curtailment_df : pd.DataFrame
        DataFrame detallado de renovables con columnas tipo:
        PV_ES0 0_available, PV_ES0 0_used, PV_ES0 0_curtailment,
        Wind_ES0 18_available, Wind_ES0 18_used, Wind_ES0 18_curtailment, etc.

    top_n_nodes : int | None
        Número de nodos con mayor curtailment a mostrar.

    Returns
    -------
    fig : matplotlib.figure.Figure | None
        Figura generada o None si no aplica.
    """

    params = df_SYS_settings["SYSTEM PARAMETERS"]

    horizon = params.get("Static / Multiperiod", None)

    if horizon != "Multiperiod":
        return None

    resolution = params.get("Graph resolution", "Auto")
    simulation_days = params.get("Simulation duration (days)", None)

    if simulation_days is None:
        return None

    simulation_days = int(simulation_days)

    if resolution == "Auto":
        if simulation_days >= 200:
            fig_curtailment_heatmap = plot_node_curtailment_heatmap_weekly(
                curtailment_df,
                top_n_nodes=top_n_nodes,
            )

        elif 60 <= simulation_days < 200:
            fig_curtailment_heatmap = plot_node_curtailment_heatmap_daily(
                curtailment_df,
                top_n_nodes=top_n_nodes,
            )

        else:
            fig_curtailment_heatmap = plot_node_curtailment_heatmap_hourly(
                curtailment_df,
                top_n_nodes=top_n_nodes,
            )

    elif resolution == "Hourly":
        fig_curtailment_heatmap = plot_node_curtailment_heatmap_hourly(
            curtailment_df,
            top_n_nodes=top_n_nodes,
        )

    elif resolution == "Daily":
        fig_curtailment_heatmap = plot_node_curtailment_heatmap_daily(
            curtailment_df,
            top_n_nodes=top_n_nodes,
        )

    elif resolution == "Weekly":
        fig_curtailment_heatmap = plot_node_curtailment_heatmap_weekly(
            curtailment_df,
            top_n_nodes=top_n_nodes,
        )

    else:
        return None

    return fig_curtailment_heatmap
