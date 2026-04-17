import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

def plot_dispatch_figure_weekly_average(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del balance energético diario apilado
    en MWh/día a partir de una serie horaria en MW.
    """

    if horizon != "Multiperiod":
        return None

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Nos quedamos solo con las columnas relevantes
    cols_to_plot = pos_cols + neg_cols
    dispatch_week = dispatch_clean[cols_to_plot].resample("W").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_week.index)
    for col in pos_cols:
        y = dispatch_week[col]
        ax.fill_between(
            dispatch_week.index,
            base_pos,
            base_pos + y,
            step="mid",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_week.index,
            base_pos + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_week.index)
    for col in neg_cols:
        y = dispatch_week[col]
        ax.fill_between(
            dispatch_week.index,
            base_neg,
            base_neg + y,
            step="mid",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_week.index,
            base_neg + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Weekly energy balance")
    ax.set_ylabel("Energy [MWh/week]")
    ax.set_xlabel("Time")

    # Formato eje X para datos diarios
    n_weeks = len(dispatch_week)
    
    ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=int(13/127*n_weeks-136/127)+1)) #Función para autoajustar el número de ticks del eje x según el número de datos
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_dispatch_figure_daily_average(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del balance energético diario apilado
    en MWh/día a partir de una serie horaria en MW.
    """

    if horizon != "Multiperiod":
        return None

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Nos quedamos solo con las columnas relevantes
    cols_to_plot = pos_cols + neg_cols
    dispatch_daily = dispatch_clean[cols_to_plot].resample("D").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_daily.index)
    for col in pos_cols:
        y = dispatch_daily[col]
        ax.fill_between(
            dispatch_daily.index,
            base_pos,
            base_pos + y,
            step="mid",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_daily.index,
            base_pos + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_daily.index)
    for col in neg_cols:
        y = dispatch_daily[col]
        ax.fill_between(
            dispatch_daily.index,
            base_neg,
            base_neg + y,
            step="mid",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_daily.index,
            base_neg + y,
            where="mid",
            color=colors.get(col, None),
            linewidth=1
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Daily energy balance")
    ax.set_ylabel("Energy [MWh/day]")
    ax.set_xlabel("Time")

    # Formato eje X para datos diarios
    n_days = len(dispatch_daily)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, int(7/130 * n_days))))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_dispatch_figure_hourly_snapshots(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):

    """
    Devuelve la figura del dispatch apilado.
    """

    if horizon != "Multiperiod":
        return None

    fig, ax = plt.subplots(figsize=(14, 6))

    pos_cols = ["Dispatch", "PV", "Wind", "battery_discharge", "Grid_import"]
    neg_cols = ["battery_charge", "Grid_export"]

    pos_cols = [c for c in pos_cols if c in dispatch_clean.columns]
    neg_cols = [c for c in neg_cols if c in dispatch_clean.columns]

    colors = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "battery_discharge": "#66BB6A",
        "Dispatch": "#E57373",
        "Grid_import": "#B0BEC5",
        "battery_charge": "#5C6BC0",
        "Grid_export": "#424242"
    }

    # Positivos
    base_pos = pd.Series(0.0, index=dispatch_clean.index)
    for col in pos_cols:
        y = dispatch_clean[col]
        ax.fill_between(
            dispatch_clean.index,
            base_pos,
            base_pos + y,
            step="post",
            alpha=0.9,
            label=col,
            color=colors.get(col, None)
        )
        ax.step(
            dispatch_clean.index,
            base_pos + y,
            where="post",
            color=colors.get(col, None),
            linewidth=1
        )
        base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=dispatch_clean.index)
    for col in neg_cols:
        y = dispatch_clean[col]
        ax.fill_between(
            dispatch_clean.index,
            base_neg,
            base_neg + y,
            step="post",
            alpha=0.8,
            label=col,
            color=colors.get(col, None)
        )
        base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Dispatch")
    ax.set_ylabel("Power [MW]")
    ax.set_xlabel("Time")

    # Formato eje X
    n_snapshots = len(dispatch_clean)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=int(1/15*n_snapshots-1/5)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Hh\n%d %b"))

    else:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=int(1/408*n_snapshots+9/17)))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def dispatch_graph_resolution_choice(df_SYS_settings: pd.DataFrame, dispatch_clean: pd.DataFrame):
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_snapshots = params["Simulation duration (days)"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            return plot_dispatch_figure_weekly_average(dispatch_clean, horizon)
        elif 60 <= n_snapshots < 200:
            return plot_dispatch_figure_daily_average(dispatch_clean, horizon)
        else:
            return plot_dispatch_figure_hourly_snapshots(dispatch_clean, horizon)

    elif resolution == "Hourly":
        return plot_dispatch_figure_hourly_snapshots(dispatch_clean, horizon)

    elif resolution == "Daily":
        return plot_dispatch_figure_daily_average(dispatch_clean, horizon)

    elif resolution == "Weekly":
        return plot_dispatch_figure_weekly_average(dispatch_clean, horizon)

    return None
