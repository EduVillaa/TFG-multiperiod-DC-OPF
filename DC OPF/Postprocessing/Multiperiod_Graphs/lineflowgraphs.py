import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import matplotlib.colors as mcolors
import pandas as pd

# Gráficos que muestra la línea más cargada en cada momento
 
def plot_max_line_loading_hourly_with_dominant_line(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con el máximo loading (%) en cada instante entre todas las líneas.

    Representación exacta por intervalos horarios:
    - cada escalón corresponde a una hora
    - el color del escalón indica qué línea es la dominante en esa hora
    - el área bajo cada escalón usa el mismo color
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] por línea
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")

    if line_loading.empty:
        return None

    max_loading = line_loading.max(axis=1)
    dominant_line = line_loading.idxmax(axis=1)

    if max_loading.empty:
        return None

    # Colores bien diferenciados
    unique_lines = list(line_loading.columns)
    n = len(unique_lines)

    if n <= 10:
        colors = plt.get_cmap("tab10").colors[:n]
    else:
        hues = np.linspace(0, 1, n, endpoint=False)
        colors = [mcolors.hsv_to_rgb((h, 0.85, 0.9)) for h in hues]

    color_map = dict(zip(unique_lines, colors))

    fig, ax = plt.subplots(figsize=(13, 5))

    x = max_loading.index
    y = max_loading.values
    dom = dominant_line.values

    if len(x) < 2:
        return None

    # Ancho típico del paso temporal
    dt = x[1] - x[0]

    # Para leyenda sin repetir
    already_labeled = set()

    # Dibujar cada intervalo [x[i], x[i+1]) con el color de la línea dominante en i
    for i in range(len(x) - 1):
        line_name = dom[i]
        color = color_map[line_name]

        x0 = x[i]
        x1 = x[i] + dt
        yi = y[i]

        label = line_name if line_name not in already_labeled else None
        if label is not None:
            already_labeled.add(line_name)

        # Área exacta del intervalo
        ax.fill_between(
            [x0, x1],
            [yi, yi],
            [0, 0],
            color=color,
            alpha=0.22,
            step="post"
        )

        # Línea exacta del intervalo
        ax.step(
            [x0, x1],
            [yi, yi],
            where="post",
            color=color,
            linewidth=2.8,
            label=label
        )

    # Opcional: unir visualmente todos los puntos con una línea gris fina de fondo
    ax.step(
        list(x) + [x[-1] + dt],
        list(y) + [y[-1]],
        where="post",
        color="0.75",
        linewidth=1.2,
        zorder=0
    )

    ax.axhline(
        y=100,
        color="red",
        linestyle="--",
        linewidth=1.5,
        label="Thermal limit (100%)"
    )

    ax.set_title("Maximum line loading over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Loading [%]")
    ax.grid(True, alpha=0.3)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.legend(
    title="Dominant line",
    loc="center left",
    bbox_to_anchor=(1.02, 0.5),
    borderaxespad=0
    )

    fig.tight_layout(rect=[0, 0, 0.82, 1])

    return fig

def plot_max_line_loading_daily_with_dominant_line(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con el máximo loading diario (%) entre todas las líneas.

    Para cada día:
    - se calcula el máximo loading de cada línea
    - se toma el mayor entre todas las líneas
    - se identifica qué línea fue la dominante ese día

    La representación es escalonada por días:
    - cada escalón corresponde a un día
    - el color indica qué línea fue la más congestionada ese día
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] horario por línea
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")

    if line_loading.empty:
        return None

    # Máximo diario de cada línea
    line_loading_daily = line_loading.resample("D").max()

    if line_loading_daily.empty:
        return None

    # Máximo diario global y línea dominante de cada día
    max_loading_daily = line_loading_daily.max(axis=1)
    dominant_line_daily = line_loading_daily.idxmax(axis=1)

    if max_loading_daily.empty:
        return None

    # Colores bien diferenciados
    unique_lines = list(line_loading_daily.columns)
    n = len(unique_lines)

    if n <= 10:
        colors = plt.get_cmap("tab10").colors[:n]
    else:
        hues = np.linspace(0, 1, n, endpoint=False)
        colors = [mcolors.hsv_to_rgb((h, 0.85, 0.9)) for h in hues]

    color_map = dict(zip(unique_lines, colors))

    fig, ax = plt.subplots(figsize=(13, 5))

    x = max_loading_daily.index
    y = max_loading_daily.values
    dom = dominant_line_daily.values

    if len(x) == 0:
        return None

    # Para leyenda sin repetir
    already_labeled = set()

    # Cada escalón representa un día completo
    one_day = np.timedelta64(1, "D")

    for i in range(len(x)):
        line_name = dom[i]
        color = color_map[line_name]

        x0 = x[i]
        x1 = x[i] + one_day
        yi = y[i]

        label = line_name if line_name not in already_labeled else None
        if label is not None:
            already_labeled.add(line_name)

        ax.fill_between(
            [x0, x1],
            [yi, yi],
            [0, 0],
            color=color,
            alpha=0.22,
            step="post"
        )

        ax.step(
            [x0, x1],
            [yi, yi],
            where="post",
            color=color,
            linewidth=2.8,
            label=label
        )

    # Línea gris de fondo continua
    ax.step(
        list(x) + [x[-1] + one_day],
        list(y) + [y[-1]],
        where="post",
        color="0.75",
        linewidth=1.2,
        zorder=0
    )

    ax.axhline(
        y=100,
        color="red",
        linestyle="--",
        linewidth=1.5,
        label="Thermal limit (100%)"
    )

    ax.set_title("Maximum daily line loading over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Loading [%]")
    ax.grid(True, alpha=0.3)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.legend(
    title="Dominant line",
    loc="center left",
    bbox_to_anchor=(1.02, 0.5),
    borderaxespad=0
    )

    fig.tight_layout(rect=[0, 0, 0.82, 1])

    return fig

def plot_max_line_loading_weekly_with_dominant_line(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con el máximo loading semanal (%) entre todas las líneas.

    Para cada semana:
    - se calcula el máximo loading de cada línea
    - se toma el mayor entre todas las líneas
    - se identifica qué línea fue la dominante esa semana

    La representación es escalonada por semanas:
    - cada escalón corresponde a una semana
    - el color indica qué línea fue la más congestionada esa semana
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] horario por línea
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")

    if line_loading.empty:
        return None

    # Máximo semanal de cada línea
    line_loading_weekly = line_loading.resample("W").max()

    if line_loading_weekly.empty:
        return None

    # Máximo semanal global y línea dominante de cada semana
    max_loading_weekly = line_loading_weekly.max(axis=1)
    dominant_line_weekly = line_loading_weekly.idxmax(axis=1)

    if max_loading_weekly.empty:
        return None

    # Colores bien diferenciados
    unique_lines = list(line_loading_weekly.columns)
    n = len(unique_lines)

    if n <= 10:
        colors = plt.get_cmap("tab10").colors[:n]
    else:
        hues = np.linspace(0, 1, n, endpoint=False)
        colors = [mcolors.hsv_to_rgb((h, 0.85, 0.9)) for h in hues]

    color_map = dict(zip(unique_lines, colors))

    fig, ax = plt.subplots(figsize=(13, 5))

    x = max_loading_weekly.index
    y = max_loading_weekly.values
    dom = dominant_line_weekly.values

    if len(x) == 0:
        return None

    # Para leyenda sin repetir
    already_labeled = set()

    one_week = np.timedelta64(7, "D")

    for i in range(len(x)):
        line_name = dom[i]
        color = color_map[line_name]

        x0 = x[i]
        x1 = x[i] + one_week
        yi = y[i]

        label = line_name if line_name not in already_labeled else None
        if label is not None:
            already_labeled.add(line_name)

        ax.fill_between(
            [x0, x1],
            [yi, yi],
            [0, 0],
            color=color,
            alpha=0.22,
            step="post"
        )

        ax.step(
            [x0, x1],
            [yi, yi],
            where="post",
            color=color,
            linewidth=2.8,
            label=label
        )

    # Línea gris de fondo continua
    ax.step(
        list(x) + [x[-1] + one_week],
        list(y) + [y[-1]],
        where="post",
        color="0.75",
        linewidth=1.2,
        zorder=0
    )

    ax.axhline(
        y=100,
        color="red",
        linestyle="--",
        linewidth=1.5,
        label="Thermal limit (100%)"
    )

    ax.set_title("Maximum weekly line loading over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Loading [%]")
    ax.grid(True, alpha=0.3)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.legend(
    title="Dominant line",
    loc="center left",
    bbox_to_anchor=(1.02, 0.5),
    borderaxespad=0
    )

    fig.tight_layout(rect=[0, 0, 0.82, 1])

    return fig

# Gráficos que muestran un heatmap con las líneas con mayor carga máxima

def plot_line_loading_hourly_heatmap(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "max", #con sort_by="max" muesteas las líneas con mayor loading máximo. Con sort_by="mean" muestras las líneas con mayor loading medio.
    top_n_lines: int | None = 15
):
    """
    Devuelve un heatmap del loading [%] de las líneas a lo largo del tiempo.

    Parameters
    ----------
    grid : pypsa.Network
        Red PyPSA.
    horizon : str, optional
        Solo genera figura si horizon == "Multiperiod".
    sort_by : str, optional
        Criterio para ordenar las líneas en el eje Y:
        - "max"  : de mayor a menor loading máximo
        - "mean" : de mayor a menor loading medio
        - "name" : orden alfabético
    top_n_lines : int | None, optional
        Número máximo de líneas a mostrar en el heatmap.
        Si es None, se muestran todas las líneas.

    Returns
    -------
    fig : matplotlib.figure.Figure | None
        Figura generada o None si no hay datos.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] por línea y por instante
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")
    line_loading = line_loading.loc[:, line_loading.max() > 1]

    if line_loading.empty:
        return None

    # Ordenar líneas
    if sort_by == "max":
        ordered_cols = line_loading.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = line_loading.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(line_loading.columns)
    else:
        ordered_cols = line_loading.columns

    line_loading = line_loading[ordered_cols]

    # Quedarse solo con las N líneas más cargadas
    if top_n_lines is not None:
        line_loading = line_loading.iloc[:, :top_n_lines]

    if line_loading.empty:
        return None

    # Matriz para el heatmap
    data = line_loading.T.values

    # Fechas convertidas a número para extent
    x_num = mdates.date2num(line_loading.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    dt = np.median(np.diff(x_num))

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(line_loading.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(line_loading.columns) - 0.5],
        interpolation="nearest",
        vmin=0
    )

    ax.set_yticks(np.arange(len(line_loading.columns)))
    ax.set_yticklabels(line_loading.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.set_title(f"Line loading heatmap over time - Top {len(line_loading.columns)} lines")
    ax.set_xlabel("Time")
    ax.set_ylabel("Lines")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Loading [%]")

    fig.tight_layout()

    return fig

def plot_line_loading_heatmap_daily(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "max", #con sort_by="max" muesteas las líneas con mayor loading máximo. Con sort_by="mean" muestras las líneas con mayor loading medio.
    top_n_lines: int | None = 15
):
    """
    Devuelve un heatmap del máximo loading diario [%] de las líneas.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] horario por línea
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")
    line_loading = line_loading.loc[:, line_loading.max() > 1]

    if line_loading.empty:
        return None

    # Máximo diario por línea
    line_loading_daily = line_loading.resample("D").max()

    if line_loading_daily.empty:
        return None

    # Ordenar líneas
    if sort_by == "max":
        ordered_cols = line_loading_daily.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = line_loading_daily.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(line_loading_daily.columns)
    else:
        ordered_cols = line_loading_daily.columns

    line_loading_daily = line_loading_daily[ordered_cols]

    # Quedarse solo con las N líneas más cargadas
    if top_n_lines is not None:
        line_loading_daily = line_loading_daily.iloc[:, :top_n_lines]

    if line_loading_daily.empty:
        return None

    data = line_loading_daily.T.values

    x_num = mdates.date2num(line_loading_daily.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    dt = 1.0

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(line_loading_daily.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(line_loading_daily.columns) - 0.5],
        interpolation="nearest",
        vmin=0,
        vmax=100
    )

    ax.set_yticks(np.arange(len(line_loading_daily.columns)))
    ax.set_yticklabels(line_loading_daily.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.set_title(f"Daily maximum line loading heatmap - Top {len(line_loading_daily.columns)} lines")
    ax.set_xlabel("Time")
    ax.set_ylabel("Lines")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Loading [%]")

    fig.tight_layout()

    return fig

def plot_line_loading_heatmap_weekly(
    grid,
    horizon: str = "Multiperiod",
    sort_by: str = "max", #con sort_by="max" muesteas las líneas con mayor loading máximo. Con sort_by="mean" muestras las líneas con mayor loading medio.
    top_n_lines: int | None = 15
):
    """
    Devuelve un heatmap del máximo loading semanal [%] de las líneas.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%] horario por línea
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")
    line_loading = line_loading.loc[:, line_loading.max() > 1]

    if line_loading.empty:
        return None

    # Máximo semanal por línea
    line_loading_weekly = line_loading.resample("W").max()

    if line_loading_weekly.empty:
        return None

    # Ordenar líneas
    if sort_by == "max":
        ordered_cols = line_loading_weekly.max(axis=0).sort_values(ascending=False).index
    elif sort_by == "mean":
        ordered_cols = line_loading_weekly.mean(axis=0).sort_values(ascending=False).index
    elif sort_by == "name":
        ordered_cols = sorted(line_loading_weekly.columns)
    else:
        ordered_cols = line_loading_weekly.columns

    line_loading_weekly = line_loading_weekly[ordered_cols]

    # Quedarse solo con las N líneas más cargadas
    if top_n_lines is not None:
        line_loading_weekly = line_loading_weekly.iloc[:, :top_n_lines]

    if line_loading_weekly.empty:
        return None

    data = line_loading_weekly.T.values

    x_num = mdates.date2num(line_loading_weekly.index.to_pydatetime())

    if len(x_num) < 2:
        return None

    dt = 7.0

    x_min = x_num[0] - dt / 2
    x_max = x_num[-1] + dt / 2

    fig_height = max(4, 0.45 * len(line_loading_weekly.columns))
    fig, ax = plt.subplots(figsize=(13, fig_height))

    im = ax.imshow(
        data,
        aspect="auto",
        origin="lower",
        extent=[x_min, x_max, -0.5, len(line_loading_weekly.columns) - 0.5],
        interpolation="nearest",
        vmin=0,
        vmax=100
    )

    ax.set_yticks(np.arange(len(line_loading_weekly.columns)))
    ax.set_yticklabels(line_loading_weekly.columns)

    locator = mdates.AutoDateLocator()
    ax.xaxis.set_major_locator(locator)
    ax.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator))
    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.set_title(f"Weekly maximum line loading heatmap - Top {len(line_loading_weekly.columns)} lines")
    ax.set_xlabel("Time")
    ax.set_ylabel("Lines")

    cbar = fig.colorbar(im, ax=ax)
    cbar.set_label("Loading [%]")

    fig.tight_layout()

    return fig

# Gráficos que muestran el porcentaje de carga de las líneas más cargadas. Para la versión daily y weekly se muestra 
# la media diaria y semanal del porcentaje de carga respectivamente de las líneas más cargadas

def plot_line_loading_percent_hourly(
    grid,
    horizon: str = "Multiperiod",
    top_n_lines: int = 1
):
    """
    Devuelve una figura con el % de carga de las líneas más cargadas
    respecto a su thermal limit.

    100% significa que la línea alcanza su límite térmico.

    Parameters
    ----------
    grid : pypsa.Network
        Red PyPSA.
    horizon : str, optional
        Solo genera figura si horizon == "Multiperiod".
    top_n_lines : int, optional
        Número de líneas más cargadas que se quieren mostrar.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # % loading por línea
    loading = grid.lines_t.p0.abs().div(s_nom, axis=1) * 100

    # Eliminar líneas sin datos válidos
    loading = loading.dropna(axis=1, how="all")

    if loading.empty:
        return None

    # Ordenar líneas por loading máximo y quedarse con las top N
    top_lines = loading.max(axis=0).sort_values(ascending=False).head(top_n_lines).index
    loading = loading[top_lines]

    fig, ax = plt.subplots(figsize=(12, 5))

    for line_name in loading.columns:
        ax.plot(loading.index, loading[line_name], label=line_name)

    ax.axhline(
        100,
        linestyle="--",
        color="red",
        linewidth=1.2,
        label="Thermal limit"
    )

    ax.set_title(f"Top {len(loading.columns)} line loadings over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Line loading [%]")

    n_snapshots = len(loading)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_line_loading_percent_daily_average(
    grid,
    horizon: str = "Multiperiod",
    top_n_lines: int = 1
):
    """
    Devuelve una figura con la media diaria del % de carga de las líneas
    más cargadas respecto a su thermal limit.

    La selección de las top N líneas se hace según el loading medio horario
    durante todo el periodo.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # % loading horario por línea
    loading = grid.lines_t.p0.abs().div(s_nom, axis=1) * 100
    loading = loading.dropna(axis=1, how="all")

    if loading.empty:
        return None

    # Seleccionar las top N líneas según loading medio horario
    top_lines = (
        loading.mean(axis=0)
        .sort_values(ascending=False)
        .head(top_n_lines)
        .index
    )

    loading = loading[top_lines]

    # Media diaria del porcentaje de carga
    loading_daily = loading.resample("D").mean()

    if loading_daily.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    for line_name in loading_daily.columns:
        ax.plot(
            loading_daily.index,
            loading_daily[line_name],
            marker="o",
            linewidth=1.8,
            label=line_name
        )

    ax.axhline(
        100,
        linestyle="--",
        color="red",
        linewidth=1.2,
        label="Thermal limit"
    )

    ax.set_title(f"Top {len(loading_daily.columns)} line loadings - daily average")
    ax.set_xlabel("Time")
    ax.set_ylabel("Average line loading [%]")

    n_days = len(loading_daily)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_days / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_line_loading_percent_weekly_average(
    grid,
    horizon: str = "Multiperiod",
    top_n_lines: int = 1
):
    """
    Devuelve una figura con la media semanal del % de carga de las líneas
    más cargadas respecto a su thermal limit.

    La selección de las top N líneas se hace según el loading medio horario
    durante todo el periodo.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # % loading horario por línea
    loading = grid.lines_t.p0.abs().div(s_nom, axis=1) * 100
    loading = loading.dropna(axis=1, how="all")

    if loading.empty:
        return None

    # Seleccionar las top N líneas según loading medio horario
    top_lines = (
        loading.mean(axis=0)
        .sort_values(ascending=False)
        .head(top_n_lines)
        .index
    )

    loading = loading[top_lines]

    # Media semanal del porcentaje de carga
    loading_weekly = loading.resample("W").mean()

    if loading_weekly.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    for line_name in loading_weekly.columns:
        ax.plot(
            loading_weekly.index,
            loading_weekly[line_name],
            marker="o",
            linewidth=1.8,
            label=line_name
        )

    ax.axhline(
        100,
        linestyle="--",
        color="red",
        linewidth=1.2,
        label="Thermal limit"
    )

    ax.set_title(f"Top {len(loading_weekly.columns)} line loadings - weekly average")
    ax.set_xlabel("Time")
    ax.set_ylabel("Average line loading [%]")

    n_weeks = len(loading_weekly)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    plt.setp(ax.get_xticklabels(), rotation=30, ha="right")

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


# Función que decide la resolución que tendrán los gráficos

def maxloading_graph_resolution_choice(df_SYS_settings: pd.DataFrame, grid):
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_snapshots = params["Simulation duration (days)"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            fig_max_line_loading = plot_max_line_loading_weekly_with_dominant_line(grid, horizon)
            fig_heatmap_line_loading = plot_line_loading_heatmap_weekly(grid, horizon, "max")
            fig_most_loaded_lines_charge = plot_line_loading_percent_weekly_average(grid, horizon)
        elif 60 <= n_snapshots < 200:
            fig_max_line_loading = plot_max_line_loading_daily_with_dominant_line(grid, horizon)
            fig_heatmap_line_loading = plot_line_loading_heatmap_daily(grid, horizon, "max")
            fig_most_loaded_lines_charge = plot_line_loading_percent_daily_average(grid, horizon)
        else:
            fig_max_line_loading = plot_max_line_loading_hourly_with_dominant_line(grid, horizon)
            fig_heatmap_line_loading = plot_line_loading_hourly_heatmap(grid, horizon, "max")
            fig_most_loaded_lines_charge = plot_line_loading_percent_hourly(grid, horizon)

    elif resolution == "Hourly":
        fig_max_line_loading = plot_max_line_loading_hourly_with_dominant_line(grid, horizon)
        fig_heatmap_line_loading = plot_line_loading_hourly_heatmap(grid, horizon, "max")
        fig_most_loaded_lines_charge = plot_line_loading_percent_hourly(grid, horizon)
    elif resolution == "Daily":
        fig_max_line_loading = plot_max_line_loading_daily_with_dominant_line(grid, horizon)
        fig_heatmap_line_loading = plot_line_loading_heatmap_daily(grid, horizon, "max")
        fig_most_loaded_lines_charge = plot_line_loading_percent_daily_average(grid, horizon)
    elif resolution == "Weekly":
        fig_max_line_loading = plot_max_line_loading_weekly_with_dominant_line(grid, horizon)
        fig_heatmap_line_loading = plot_line_loading_heatmap_weekly(grid, horizon, "max")
        fig_most_loaded_lines_charge = plot_line_loading_percent_weekly_average(grid, horizon)

    else:
        return None, None

    return fig_heatmap_line_loading, fig_max_line_loading, fig_most_loaded_lines_charge


def plot_line_loading_histogram_top_lines(
    grid,
    horizon: str = "Multiperiod",
    top_n: int = 3
):
    """
    Histograma del loading [%] para las top N líneas más críticas.

    Criterio:
    - Selección basada en el máximo loading alcanzado por cada línea
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%]
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")

    if line_loading.empty:
        return None

    # 🔹 Selección robusta: top líneas por máximo loading
    top_lines = (
        line_loading.max()
        .sort_values(ascending=False)
        .head(top_n)
        .index
    )

    bins = np.arange(0, 110, 10)

    fig, ax = plt.subplots(figsize=(10, 5))

    for line in top_lines:
        ax.hist(
            line_loading[line].dropna(),
            bins=bins,
            histtype="step",
            linewidth=2,
            label=line
        )

    ax.set_title(f"Distribution of line loading [%] (Top {top_n} lines)")
    ax.set_xlabel("Loading [%]")
    ax.set_ylabel("Number of hours")

    ax.set_xticks(bins)
    ax.grid(True, alpha=0.3)

    ax.legend()

    fig.tight_layout()

    return fig


def plot_line_loading_histogram_global(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Histograma del loading [%] agregando todas las líneas.
    """

    if horizon != "Multiperiod":
        return None

    if grid.lines.empty or grid.lines_t.p0.empty:
        return None

    if "s_nom" not in grid.lines.columns:
        return None

    s_nom = grid.lines["s_nom"].replace(0, np.nan)

    # Loading [%]
    line_loading = grid.lines_t.p0.abs().divide(s_nom, axis=1) * 100
    line_loading = line_loading.dropna(axis=1, how="all")

    if line_loading.empty:
        return None

    bins = np.arange(0, 110, 10)

    fig, ax = plt.subplots(figsize=(10, 5))

    all_values = line_loading.values.flatten()

    ax.hist(
        all_values,
        bins=bins,
        edgecolor="black",
        alpha=0.6
    )

    ax.set_title("Distribution of line loading [%] (All lines)")
    ax.set_xlabel("Loading [%]")
    ax.set_ylabel("Number of hours")

    ax.set_xticks(bins)
    ax.grid(True, alpha=0.3)

    fig.tight_layout()

    return fig