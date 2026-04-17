import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import numpy as np
import pypsa
from matplotlib.figure import Figure

def plot_renewable_share_hourly(dispatch_clean: pd.DataFrame, grid, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con el renewable share horario:
    (PV + Wind) / Load

    Se representa en p.u.:
    - 1.0 => la generación renovable iguala la carga
    - >1.0 => excedente renovable
    - <1.0 => déficit renovable
    """

    if horizon != "Multiperiod":
        return None

    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    total_renewable = pv + wind

    # Carga total horaria
    load = grid.loads_t.p.sum(axis=1)

    # Alinear índices por seguridad
    total_renewable, load = total_renewable.align(load, join="inner")

    if len(load) == 0:
        return None

    # Evitar división por cero
    renewable_share = total_renewable / load.replace(0, np.nan)

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        renewable_share.index,
        renewable_share.values,
        linewidth=1.8,
        label="Renewable share"
    )

    ax.axhline(1.0, linestyle="--", color="red", linewidth=1.2, label="100% of load")

    ax.set_title("Hourly renewable share")
    ax.set_xlabel("Time")
    ax.set_ylabel("Renewable share [p.u.]")

    n_snapshots = len(renewable_share)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend()
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig



def plot_renewable_share_daily(dispatch_clean: pd.DataFrame, grid, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con el renewable share diario:
    (energía renovable diaria) / (energía de carga diaria)

    Se representa en p.u.:
    - 1.0 => renovables cubren toda la demanda del día
    - >1.0 => excedente renovable
    - <1.0 => déficit renovable
    """

    if horizon != "Multiperiod":
        return None

    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    total_renewable = pv + wind

    # Carga total horaria
    load = grid.loads_t.p.sum(axis=1)

    # Alinear índices
    total_renewable, load = total_renewable.align(load, join="inner")

    if len(load) == 0:
        return None

    # Energía diaria
    renewable_daily = total_renewable.resample("D").sum()
    load_daily = load.resample("D").sum()

    # Evitar división por cero
    renewable_share_daily = renewable_daily / load_daily.replace(0, np.nan)

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.step(
        renewable_share_daily.index,
        renewable_share_daily.values,
        where="mid",
        linewidth=1.8,
        label="Daily renewable share"
    )

    ax.fill_between(
        renewable_share_daily.index,
        0,
        renewable_share_daily.values,
        step="mid",
        alpha=0.25
    )

    ax.axhline(1.0, linestyle="--", color="red", linewidth=1.2, label="100% of load")

    ax.set_title("Daily renewable share")
    ax.set_xlabel("Time")
    ax.set_ylabel("Renewable share [p.u.]")

    n_days = len(renewable_share_daily)

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
    fig.tight_layout()

    return fig

def plot_renewable_share_weekly(dispatch_clean: pd.DataFrame, grid, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con el renewable share semanal:
    (energía renovable semanal) / (energía de carga semanal)

    Se representa en p.u.:
    - 1.0 => renovables cubren toda la demanda de la semana
    - >1.0 => excedente renovable
    - <1.0 => déficit renovable
    """

    if horizon != "Multiperiod":
        return None

    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    total_renewable = pv + wind

    # Carga total horaria
    load = grid.loads_t.p.sum(axis=1)

    # Alinear índices
    total_renewable, load = total_renewable.align(load, join="inner")

    if len(load) == 0:
        return None

    # Energía semanal
    renewable_weekly = total_renewable.resample("W").sum()
    load_weekly = load.resample("W").sum()

    # Evitar división por cero
    renewable_share_weekly = renewable_weekly / load_weekly.replace(0, np.nan)

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.step(
        renewable_share_weekly.index,
        renewable_share_weekly.values,
        where="mid",
        linewidth=1.8,
        label="Weekly renewable share"
    )

    ax.fill_between(
        renewable_share_weekly.index,
        0,
        renewable_share_weekly.values,
        step="mid",
        alpha=0.25
    )

    ax.axhline(1.0, linestyle="--", color="red", linewidth=1.2, label="100% of load")

    ax.set_title("Weekly renewable share")
    ax.set_xlabel("Time")
    ax.set_ylabel("Renewable share [p.u.]")

    n_weeks = len(renewable_share_weekly)

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
    fig.tight_layout()

    return fig


def renewableshare_graph_resolution_choice(
    df_SYS_settings: pd.DataFrame,
    dispatch_clean: pd.DataFrame,
    grid: pypsa.Network
) -> Figure | None:

    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_snapshots = params["Simulation duration (days)"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            return plot_renewable_share_weekly(dispatch_clean, grid, horizon)
        elif 60 <= n_snapshots < 200:
            return plot_renewable_share_daily(dispatch_clean, grid, horizon)
        else:
            return plot_renewable_share_hourly(dispatch_clean, grid, horizon)

    elif resolution == "Hourly":
        return plot_renewable_share_hourly(dispatch_clean, grid, horizon)

    elif resolution == "Daily":
        return plot_renewable_share_daily(dispatch_clean, grid, horizon)

    elif resolution == "Weekly":
        return plot_renewable_share_weekly(dispatch_clean, grid, horizon)

    return None