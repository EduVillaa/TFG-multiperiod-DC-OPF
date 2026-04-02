import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import pypsa

def plot_total_soc_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total de todas las baterías modeladas como Store + Link.
    """

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if len(store_cols) == 0:
        return None

    soc_total = grid.stores_t.e[store_cols].sum(axis=1)

    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(soc_total.index, soc_total.values, label="Total SOC")

    capacity = pd.to_numeric(
        grid.stores.loc[store_cols, "e_nom"], errors="coerce"
    ).fillna(0).sum()

    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC")

    n_snapshots = len(soc_total)

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

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_soc_per_battery_figure(grid, horizon: str = "Multiperiod"):
    
    """
    Devuelve la figura del SOC separado por baterías
    modeladas como Store + Link.
    """

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if not store_cols:
        return None

    soc = grid.stores_t.e[store_cols].copy()

    if soc.empty:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    for battery_name in soc.columns:
        # --- CAPACIDAD CORRECTA ---
        if "e_nom_opt" in grid.stores.columns and pd.notna(grid.stores.loc[battery_name, "e_nom_opt"]):
            capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom_opt"], errors="coerce")
        else:
            capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom"], errors="coerce")

        ax.plot(
            soc.index,
            soc[battery_name],
            label=f"{battery_name} ({capacity:.1f} MWh)"
        )

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Battery SOC by unit")

    n_snapshots = len(soc)

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

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def plot_soc_per_battery_daily_average_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC medio diario separado por baterías
    modeladas como Store + Link.
    """

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if len(store_cols) == 0:
        return None

    soc = grid.stores_t.e[store_cols].copy()
    if soc.empty:
        return None

    soc_daily = soc.resample("D").mean()

    fig, ax = plt.subplots(figsize=(12, 5))

    for battery_name in soc_daily.columns:
            if "e_nom_opt" in grid.stores.columns and pd.notna(grid.stores.loc[battery_name, "e_nom_opt"]):
                capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom_opt"], errors="coerce")
            else:
                capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom"], errors="coerce")

            ax.plot(
                soc_daily.index,
                soc_daily[battery_name],
                label=f"{battery_name} ({capacity:.1f} MWh)"
            )

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Battery SOC by unit (daily average)")

    n_days = len(soc_daily)

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

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_total_soc_daily_stats_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total con:
    - media diaria
    - mínimo diario
    - máximo diario
    - banda sombreada min-max
    para baterías modeladas como Store + Link
    """

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if len(store_cols) == 0:
        return None

    soc = grid.stores_t.e[store_cols].copy()
    soc_total = soc.sum(axis=1)

    soc_daily_mean = soc_total.resample("D").mean()
    soc_daily_min = soc_total.resample("D").min()
    soc_daily_max = soc_total.resample("D").max()

    fig, ax = plt.subplots(figsize=(10, 4))

    # Banda min-max
    ax.fill_between(
        soc_daily_mean.index,
        soc_daily_min.values,
        soc_daily_max.values,
        alpha=0.25,
        label="Daily min-max range"
    )

    # Curvas
    ax.plot(soc_daily_mean.index, soc_daily_mean.values, linewidth=2, label="Daily mean SOC")
    ax.plot(soc_daily_min.index, soc_daily_min.values, linestyle="--", linewidth=1.2, label="Daily minimum SOC")
    ax.plot(soc_daily_max.index, soc_daily_max.values, linestyle="--", linewidth=1.2, label="Daily maximum SOC")

    # Capacidad total
    capacity = grid.stores.loc[store_cols, "e_nom"].sum()
    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC (daily statistics)")

    n_days = len(soc_daily_mean)

    if n_days <= 14:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_days <= 90:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    else:
        interval = max(1, int(n_days / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_total_soc_weekly_stats_figure(grid, horizon: str = "Multiperiod"):
    """
    Devuelve la figura del SOC total con:
    - media semanal
    - mínimo semanal
    - máximo semanal
    - banda sombreada min-max
    para baterías modeladas como Store + Link
    """

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if len(store_cols) == 0:
        return None

    soc = grid.stores_t.e[store_cols].copy()
    soc_total = soc.sum(axis=1)

    soc_weekly_mean = soc_total.resample("W").mean()
    soc_weekly_min = soc_total.resample("W").min()
    soc_weekly_max = soc_total.resample("W").max()

    fig, ax = plt.subplots(figsize=(10, 4))

    # Banda min-max
    ax.fill_between(
        soc_weekly_mean.index,
        soc_weekly_min.values,
        soc_weekly_max.values,
        alpha=0.25,
        label="Weekly min-max range"
    )

    # Curvas
    ax.plot(soc_weekly_mean.index, soc_weekly_mean.values, linewidth=2, label="Weekly mean SOC")
    ax.plot(soc_weekly_min.index, soc_weekly_min.values, linestyle="--", linewidth=1.2, label="Weekly minimum SOC")
    ax.plot(soc_weekly_max.index, soc_weekly_max.values, linestyle="--", linewidth=1.2, label="Weekly maximum SOC")

    # Capacidad total
    capacity = grid.stores.loc[store_cols, "e_nom"].sum()
    ax.axhline(y=capacity, linestyle="--", color="red", label="Max capacity")

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Total battery SOC (weekly statistics)")

    n_weeks = len(soc_weekly_mean)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def plot_soc_per_battery_weekly_average_figure(grid, horizon: str = "Multiperiod"):

    if horizon != "Multiperiod":
        return None

    store_cols = [c for c in grid.stores_t.e.columns if c.startswith("BatteryStore_")]
    if len(store_cols) == 0:
        return None

    soc = grid.stores_t.e[store_cols].copy()

    if soc.empty:
        return None

    # Media semanal por batería
    soc_weekly = soc.resample("W").mean()

    fig, ax = plt.subplots(figsize=(12, 5))

    for battery_name in soc_weekly.columns:
        if "e_nom_opt" in grid.stores.columns and pd.notna(grid.stores.loc[battery_name, "e_nom_opt"]):
            capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom_opt"], errors="coerce")
        else:
            capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom"], errors="coerce")

        ax.plot(
            soc_weekly.index,
            soc_weekly[battery_name],
            label=f"{battery_name} ({capacity:.1f} MWh)"
        )

    # Línea de capacidad máxima por batería
    for battery_name in soc_weekly.columns:
        if battery_name in grid.stores.index:
            if "e_nom_opt" in grid.stores.columns and pd.notna(grid.stores.loc[battery_name, "e_nom_opt"]):
                capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom_opt"], errors="coerce")
            else:
                capacity = pd.to_numeric(grid.stores.loc[battery_name, "e_nom"], errors="coerce")

            ax.axhline(
                y=capacity,
                linestyle="--",
                linewidth=1,
                alpha=0.5
            )

    ax.set_xlabel("Time")
    ax.set_ylabel("State of charge [MWh]")
    ax.set_title("Battery SOC by unit (weekly average)")

    n_weeks = len(soc_weekly)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def SOC_graph_resolution_choice(df_SYS_settings: pd.DataFrame, grid: pypsa.Network):
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]

    if df_SYS_settings.loc[7, "SYSTEM PARAMETERS"] == "Auto":
        if df_SYS_settings.loc[6, "SYSTEM PARAMETERS"] >= 200:
            fig_soc_batteries = plot_soc_per_battery_weekly_average_figure(grid, horizon)
            fig_soc_total = plot_total_soc_weekly_stats_figure(grid, horizon)
        elif 60 <= df_SYS_settings.loc[6, "SYSTEM PARAMETERS"] < 200:
            fig_soc_batteries = plot_soc_per_battery_daily_average_figure(grid, horizon)
            fig_soc_total = plot_total_soc_daily_stats_figure(grid, horizon)
        else:
            fig_soc_batteries = plot_soc_per_battery_figure(grid, horizon)
            fig_soc_total = plot_total_soc_figure(grid, horizon)

    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"] == "Hourly":
        fig_soc_batteries = plot_soc_per_battery_figure(grid, horizon)
        fig_soc_total = plot_total_soc_figure(grid, horizon)

    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"] == "Daily":
        fig_soc_batteries = plot_soc_per_battery_daily_average_figure(grid, horizon)
        fig_soc_total = plot_total_soc_daily_stats_figure(grid, horizon)

    elif df_SYS_settings.loc[7, "SYSTEM PARAMETERS"] == "Weekly":
        fig_soc_batteries = plot_soc_per_battery_weekly_average_figure(grid, horizon)
        fig_soc_total = plot_total_soc_weekly_stats_figure(grid, horizon)

    else:
        return None, None

    return fig_soc_total, fig_soc_batteries