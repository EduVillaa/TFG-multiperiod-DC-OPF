import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pandas as pd

def plot_total_load_hourly(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con la carga total por hora [MW]
    """

    if horizon != "Multiperiod":
        return None

    # Carga total
    total_load = grid.loads_t.p.sum(axis=1)

    if total_load.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        total_load.index,
        total_load.values,
        linewidth=1.8,
        label="Total load"
    )

    ax.set_title("Total load")
    ax.set_xlabel("Time")
    ax.set_ylabel("Power [MW]")

    # Formato eje X adaptativo
    n_snapshots = len(total_load)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 21:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend()
    ax.grid(True, axis="y")
    ax.fill_between(
    total_load.index,
    0,
    total_load.values,
    alpha=0.2
)
    fig.tight_layout()

    return fig

def plot_total_load_daily_energy(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con la energía total diaria [MWh/day]
    """

    if horizon != "Multiperiod":
        return None

    # Carga total
    total_load = grid.loads_t.p.sum(axis=1)

    if total_load.empty:
        return None

    # Energía diaria (asumiendo datos horarios)
    load_daily = total_load.resample("D").sum()

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.step(
        load_daily.index,
        load_daily.values,
        where="mid",
        linewidth=1.8,
        label="Daily load"
    )

    ax.set_title("Daily total load")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/day]")

    # Formato eje X
    n_days = len(load_daily)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_days / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")

    ax.fill_between(
        load_daily.index,
        0,
        load_daily.values,
        step="mid",
        alpha=0.2
    )

    fig.tight_layout()

    return fig


def plot_total_load_weekly_energy(
    grid,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con la energía total semanal [MWh/week]
    """

    if horizon != "Multiperiod":
        return None

    # Carga total
    total_load = grid.loads_t.p.sum(axis=1)

    if total_load.empty:
        return None

    # Energía semanal
    load_weekly = total_load.resample("W").sum()

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.step(
        load_weekly.index,
        load_weekly.values,
        where="mid",
        linewidth=1.8,
        label="Weekly load"
    )

    ax.set_title("Weekly total load")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/week]")

    # Formato eje X
    n_weeks = len(load_weekly)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")

    ax.fill_between(
        load_weekly.index,
        0,
        load_weekly.values,
        step="mid",
        alpha=0.2
    )

    fig.tight_layout()

    return fig


def total_load_graph_resolution_choice(df_SYS_settings: pd.DataFrame, grid):
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_days = params["Simulation duration (days)"]

    if resolution == "Auto":
        if n_days >= 200:
            return plot_total_load_weekly_energy(grid, horizon)
        elif 60 <= n_days < 200:
            return plot_total_load_daily_energy(grid, horizon)
        else:
            return plot_total_load_hourly(grid, horizon)

    elif resolution == "Hourly":
        return plot_total_load_hourly(grid, horizon)

    elif resolution == "Daily":
        return plot_total_load_daily_energy(grid, horizon)

    elif resolution == "Weekly":
        return plot_total_load_weekly_energy(grid, horizon)

    return None
