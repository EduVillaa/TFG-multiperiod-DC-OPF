import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.figure import Figure
import numpy as np

def meanprices_hourly(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con tres curvas de precios [€/MWh]:
    - Precio medio del sistema
    - Precio mínimo entre buses
    - Precio máximo entre buses
    """

    if horizon != "Multiperiod":
        return None

    # Precios nodales
    prices = grid.buses_t.marginal_price
    
    if prices is None or prices.empty:
        return None

    # Por seguridad, eliminar columnas totalmente vacías
    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None
    
    #Quitamos del dataframe la columna de precios marginales en el PCC y en los buses de las baterías.

    prices = prices[
    [
        col for col in prices.columns
        if col != "PCC" and "battery" not in col.lower()
    ]
]
    # Estadísticos por instante temporal
    mean_price = prices.mean(axis=1)
    min_price = prices.min(axis=1)
    max_price = prices.max(axis=1)

    if mean_price.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        mean_price.index,
        mean_price.values,
        linewidth=2.0,
        label="Mean price"
    )

    ax.plot(
        min_price.index,
        min_price.values,
        linewidth=1.5,
        linestyle="--",
        label="Min nodal price"
    )

    ax.plot(
        max_price.index,
        max_price.values,
        linewidth=1.5,
        linestyle="--",
        label="Max nodal price"
    )

    ax.set_title("Bus marginal prices")
    ax.set_xlabel("Time")
    ax.set_ylabel("Price [€/MWh]")

    ax.fill_between(
    mean_price.index,
    min_price.values,
    max_price.values,
    alpha=0.15,
    label="Nodal price range"
)

    # Formato eje X adaptativo
    n_snapshots = len(mean_price)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 21:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def meanprices_daily(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con tres curvas diarias de precios [€/MWh]:
    - Precio medio diario del sistema
    - Precio mínimo diario entre buses
    - Precio máximo diario entre buses
    """

    if horizon != "Multiperiod":
        return None

    # Precios nodales
    prices = grid.buses_t.marginal_price

    if prices is None or prices.empty:
        return None

    # Eliminar columnas totalmente vacías
    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    # Quitar PCC y buses de baterías
    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if prices.empty or prices.shape[1] == 0:
        return None

    # Asegurar índice datetime
    prices = prices.copy()
    prices.index = pd.to_datetime(prices.index)

    # Estadísticos horarios entre buses
    mean_price_hourly = prices.mean(axis=1)
    min_price_hourly = prices.min(axis=1)
    max_price_hourly = prices.max(axis=1)

    if mean_price_hourly.empty:
        return None

    # Agregación diaria
    mean_price_daily = mean_price_hourly.resample("D").mean()
    min_price_daily = min_price_hourly.resample("D").min()
    max_price_daily = max_price_hourly.resample("D").max()

    if mean_price_daily.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        mean_price_daily.index,
        mean_price_daily.values,
        linewidth=2.0,
        label="Mean daily price"
    )

    ax.plot(
        min_price_daily.index,
        min_price_daily.values,
        linewidth=1.5,
        linestyle="--",
        label="Min daily nodal price"
    )

    ax.plot(
        max_price_daily.index,
        max_price_daily.values,
        linewidth=1.5,
        linestyle="--",
        label="Max daily nodal price"
    )

    ax.fill_between(
        mean_price_daily.index,
        min_price_daily.values,
        max_price_daily.values,
        alpha=0.15,
        label="Daily nodal price range"
    )

    ax.set_title("Daily bus marginal prices")
    ax.set_xlabel("Time")
    ax.set_ylabel("Price [€/MWh]")

    # Formato adaptativo del eje X
    n_days = len(mean_price_daily)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=7))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_days / 12))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def meanprices_weekly(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con tres curvas semanales de precios [€/MWh]:
    - Precio medio semanal del sistema
    - Precio mínimo semanal entre buses
    - Precio máximo semanal entre buses
    """

    if horizon != "Multiperiod":
        return None

    # Precios nodales
    prices = grid.buses_t.marginal_price

    if prices is None or prices.empty:
        return None

    # Eliminar columnas totalmente vacías
    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    # Quitar PCC y buses de baterías
    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if prices.empty or prices.shape[1] == 0:
        return None

    # Asegurar índice datetime
    prices = prices.copy()
    prices.index = pd.to_datetime(prices.index)

    # Estadísticos horarios entre buses
    mean_price_hourly = prices.mean(axis=1)
    min_price_hourly = prices.min(axis=1)
    max_price_hourly = prices.max(axis=1)

    if mean_price_hourly.empty:
        return None

    # Agregación semanal
    mean_price_weekly = mean_price_hourly.resample("W").mean()
    min_price_weekly = min_price_hourly.resample("W").min()
    max_price_weekly = max_price_hourly.resample("W").max()

    if mean_price_weekly.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        mean_price_weekly.index,
        mean_price_weekly.values,
        linewidth=2.0,
        label="Mean weekly price"
    )

    ax.plot(
        min_price_weekly.index,
        min_price_weekly.values,
        linewidth=1.5,
        linestyle="--",
        label="Min weekly nodal price"
    )

    ax.plot(
        max_price_weekly.index,
        max_price_weekly.values,
        linewidth=1.5,
        linestyle="--",
        label="Max weekly nodal price"
    )

    ax.fill_between(
        mean_price_weekly.index,
        min_price_weekly.values,
        max_price_weekly.values,
        alpha=0.15,
        label="Weekly nodal price range"
    )

    ax.set_title("Weekly bus marginal prices")
    ax.set_xlabel("Time")
    ax.set_ylabel("Price [€/MWh]")

    # Formato adaptativo del eje X
    n_weeks = len(mean_price_weekly)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_weeks / 12))
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_nodal_price_hourly_heatmap(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "mean",
    top_n_buses: int | None = None
):
    """
    Devuelve un heatmap de precios marginales nodales [€/MWh] a lo largo del tiempo.

    Parameters
    ----------
    grid : pypsa.Network
        Red PyPSA.
    horizon : str, optional
        Solo genera figura si horizon == "Multiperiod".
    sort_by : str, optional
        Criterio para ordenar los buses en el eje Y:
        - "max"  : de mayor a menor precio máximo
        - "mean" : de mayor a menor precio medio
        - "name" : orden alfabético

    Returns
    -------
    fig : matplotlib.figure.Figure | None
        Figura generada o None si no hay datos.
    """

    if horizon != "Multiperiod":
        return None

    if not hasattr(grid, "buses_t") or grid.buses_t.marginal_price.empty:
        return None

    prices = grid.buses_t.marginal_price.copy()

    if prices is None or prices.empty:
        return None

    # Eliminar columnas totalmente vacías
    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    # Quitar PCC y buses de baterías
    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if top_n_buses is not None:
        if top_n_buses <= 0:
            return None

        top_cols = prices.mean(axis=0).sort_values(ascending=False).head(top_n_buses).index
        prices = prices[top_cols]

    if prices.empty:
        return None

    # Ordenar buses
    if sort_by == "max":
        ordered_cols = prices.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = prices.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(prices.columns)
    else:
        ordered_cols = prices.columns

    prices = prices[ordered_cols]

    # Matriz para heatmap
    data = prices.T.values   # filas = buses, columnas = tiempo

    # Fechas convertidas a número para extent
    x_num = mdates.date2num(prices.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    # Ancho típico del paso temporal
    dt = np.median(np.diff(x_num))

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(prices.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(prices.columns) - 0.5],
        interpolation="nearest"
    )

    # Eje Y con nombres de buses
    ax.set_yticks(np.arange(len(prices.columns)))
    ax.set_yticklabels(prices.columns)

    # Eje X con fechas
    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.set_title("Nodal marginal price heatmap over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Buses")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Price [€/MWh]")

    fig.tight_layout()

    return fig

def plot_nodal_price_heatmap_daily(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "mean",
    top_n_buses: int | None = None
):
    """
    Devuelve un heatmap del precio marginal medio diario [€/MWh] de los buses.
    """

    if horizon != "Multiperiod":
        return None

    if not hasattr(grid, "buses_t") or grid.buses_t.marginal_price.empty:
        return None

    prices = grid.buses_t.marginal_price.copy()

    if prices is None or prices.empty:
        return None

    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if prices.empty or prices.shape[1] == 0:
        return None

    prices.index = pd.to_datetime(prices.index)

    # Media diaria por bus
    prices_daily = prices.resample("D").mean()

    if prices_daily.empty or prices_daily.shape[1] == 0:
        return None

    # Quedarse solo con los buses con mayor precio medio diario
    if top_n_buses is not None:
        if top_n_buses <= 0:
            return None

        top_cols = (
            prices_daily.mean(axis=0)
            .sort_values(ascending=False)
            .head(top_n_buses)
            .index
        )
        prices_daily = prices_daily[top_cols]

    if prices_daily.empty or prices_daily.shape[1] == 0:
        return None

    # Ordenar buses
    if sort_by == "max":
        ordered_cols = prices_daily.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = prices_daily.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(prices_daily.columns)
    else:
        ordered_cols = prices_daily.columns

    prices_daily = prices_daily[ordered_cols]

    data = prices_daily.T.values
    x_num = mdates.date2num(prices_daily.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    dt = 1.0  # 1 día

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(prices_daily.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(prices_daily.columns) - 0.5],
        interpolation="nearest"
    )

    ax.set_yticks(np.arange(len(prices_daily.columns)))
    ax.set_yticklabels(prices_daily.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    title = "Daily mean nodal marginal price heatmap"
    if top_n_buses is not None:
        title += f" (top {len(prices_daily.columns)} buses by mean price)"
    ax.set_title(title)

    ax.set_xlabel("Time")
    ax.set_ylabel("Buses")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Price [€/MWh]")

    fig.tight_layout()

    return fig

def plot_nodal_price_heatmap_weekly(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "mean",
    top_n_buses: int | None = None
):
    """
    Devuelve un heatmap del precio marginal medio semanal [€/MWh] de los buses.
    """

    if horizon != "Multiperiod":
        return None

    if not hasattr(grid, "buses_t") or grid.buses_t.marginal_price.empty:
        return None

    prices = grid.buses_t.marginal_price.copy()

    if prices is None or prices.empty:
        return None

    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if prices.empty or prices.shape[1] == 0:
        return None

    prices.index = pd.to_datetime(prices.index)

    # Media semanal por bus
    prices_weekly = prices.resample("W").mean()

    if prices_weekly.empty or prices_weekly.shape[1] == 0:
        return None

    # Quedarse solo con los buses con mayor precio medio semanal
    if top_n_buses is not None:
        if top_n_buses <= 0:
            return None

        top_cols = (
            prices_weekly.mean(axis=0)
            .sort_values(ascending=False)
            .head(top_n_buses)
            .index
        )
        prices_weekly = prices_weekly[top_cols]

    if prices_weekly.empty or prices_weekly.shape[1] == 0:
        return None

    # Ordenar buses
    if sort_by == "max":
        ordered_cols = prices_weekly.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = prices_weekly.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(prices_weekly.columns)
    else:
        ordered_cols = prices_weekly.columns

    prices_weekly = prices_weekly[ordered_cols]

    data = prices_weekly.T.values
    x_num = mdates.date2num(prices_weekly.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    dt = 7.0  # 7 días

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(prices_weekly.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(prices_weekly.columns) - 0.5],
        interpolation="nearest"
    )

    ax.set_yticks(np.arange(len(prices_weekly.columns)))
    ax.set_yticklabels(prices_weekly.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    title = "Weekly mean nodal marginal price heatmap"
    if top_n_buses is not None:
        title += f" (top {len(prices_weekly.columns)} buses by mean price)"
    ax.set_title(title)

    ax.set_xlabel("Time")
    ax.set_ylabel("Buses")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Price [€/MWh]")

    fig.tight_layout()

    return fig

def prices_graph_resolution_choice(
    df_SYS_settings: pd.DataFrame,
    grid: pd.DataFrame, top_n_buses,
) -> tuple[Figure | None, Figure | None]:
    
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_snapshots = params["Simulation duration (days)"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            fig_meanprices = meanprices_weekly(grid, horizon)
            fig_heatmap = plot_nodal_price_heatmap_weekly(grid, horizon, "mean", top_n_buses)
        elif 60 <= n_snapshots < 200:
            fig_meanprices = meanprices_daily(grid, horizon)
            fig_heatmap = plot_nodal_price_heatmap_daily(grid, horizon, "mean", top_n_buses)
        else:
            fig_meanprices = meanprices_hourly(grid, horizon)
            fig_heatmap = plot_nodal_price_hourly_heatmap(grid, horizon, "mean", top_n_buses)
    elif resolution == "Hourly":
        fig_meanprices = meanprices_hourly(grid, horizon)
        fig_heatmap = plot_nodal_price_hourly_heatmap(grid, horizon, "mean", top_n_buses)
    elif resolution == "Daily":
        fig_meanprices = meanprices_daily(grid, horizon)
        fig_heatmap = plot_nodal_price_heatmap_daily(grid, horizon, "mean", top_n_buses)
    elif resolution == "Weekly":
        fig_meanprices = meanprices_weekly(grid, horizon)
        fig_heatmap = plot_nodal_price_heatmap_weekly(grid, horizon, "mean", top_n_buses)

    else:
        fig_meanprices, fig_heatmap = None, None

    return fig_meanprices, fig_heatmap


def nodal_price_histogram(
    grid,
    horizon: str = "Multiperiod",
    max_bins: int = 30
):
    """
    Devuelve una figura con el histograma de todos los precios nodales [€/MWh].
    El número de bins se ajusta automáticamente al rango de los datos.
    """

    if horizon != "Multiperiod":
        return None

    prices = grid.buses_t.marginal_price

    if prices is None or prices.empty:
        return None

    prices = prices.dropna(axis=1, how="all")

    if prices.empty:
        return None

    prices = prices[
        [
            col for col in prices.columns
            if col != "PCC" and "battery" not in col.lower()
        ]
    ]

    if prices.empty or prices.shape[1] == 0:
        return None

    values = prices.values.flatten()
    values = pd.Series(values).dropna()

    if values.empty:
        return None

    vmin = values.min()
    vmax = values.max()
    n = len(values)

    fig, ax = plt.subplots(figsize=(12, 5))

    # Caso: todos los valores iguales o prácticamente iguales
    if np.isclose(vmin, vmax, rtol=0, atol=1e-9):
        ax.bar([vmin], [n], width=0.1 if abs(vmin) > 1e-9 else 0.01)
        ax.set_title("Histogram of nodal marginal prices (constant value)")
        ax.set_xlabel("Price [€/MWh]")
        ax.set_ylabel("Frequency")
        ax.grid(True, axis="y")
        fig.tight_layout()
        return fig

    # Regla de Freedman–Diaconis
    q75, q25 = values.quantile([0.75, 0.25])
    iqr = q75 - q25

    if np.isclose(iqr, 0, rtol=0, atol=1e-12):
        bins = min(max_bins, values.nunique())
    else:
        bin_width = 2 * iqr / (n ** (1 / 3))

        if np.isclose(bin_width, 0, rtol=0, atol=1e-12):
            bins = min(max_bins, values.nunique())
        else:
            bins = int(np.ceil((vmax - vmin) / bin_width))

    # Seguridad extra
    bins = max(1, min(bins, max_bins))

    # Si por cualquier razón el rango es casi nulo, fuerza 1 bin
    if bins > 1 and np.isclose(vmax - vmin, 0, rtol=0, atol=1e-9):
        bins = 1

    ax.hist(values, bins=bins)

    ax.set_title("Histogram of nodal marginal prices")
    ax.set_xlabel("Price [€/MWh]")
    ax.set_ylabel("Frequency")
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig