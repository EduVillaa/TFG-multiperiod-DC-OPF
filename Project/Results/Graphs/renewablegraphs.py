import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.figure import Figure
import pypsa

def plot_total_renewable_power(
    dispatch_clean: pd.DataFrame,
    df_available_renewable: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con dos curvas por hora [MW]:
    - Potencia renovable generada total (PV + Wind)
    - Potencia renovable disponible total
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas de generación usada
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    total_renewable_used = pv + wind

    # Renovable disponible total
    if df_available_renewable is None or df_available_renewable.empty:
        return None

    total_renewable_available = df_available_renewable.sum(axis=1)

    # Alinear índices por seguridad
    common_index = total_renewable_used.index.intersection(total_renewable_available.index)

    if len(common_index) == 0:
        return None

    total_renewable_used = total_renewable_used.loc[common_index]
    total_renewable_available = total_renewable_available.loc[common_index]

    fig, ax = plt.subplots(figsize=(12, 5))

    ax.plot(
        total_renewable_used.index,
        total_renewable_used.values,
        linewidth=1.8,
        label="Total renewable generated"
    )

    ax.plot(
        total_renewable_available.index,
        total_renewable_available.values,
        linewidth=1.8,
        linestyle="--",
        label="Total renewable available"
    )

    ax.set_title("Total renewable power")
    ax.set_xlabel("Time")
    ax.set_ylabel("Power [MW]")

    # Formato eje X adaptativo
    n_snapshots = len(total_renewable_used)

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

def plot_total_renewable_daily_energy(
    dispatch_clean: pd.DataFrame,
    df_available_renewable: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con dos curvas [MWh/day]:
    - Energía renovable generada diaria (PV + Wind)
    - Energía renovable disponible diaria
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    # Potencia renovable usada
    total_renewable_used = pv + wind

    # Energía diaria usada
    renewable_daily_energy_used = total_renewable_used.resample("D").sum()

    # Renovable disponible
    if df_available_renewable is None or df_available_renewable.empty:
        return None

    total_renewable_available = df_available_renewable.sum(axis=1)

    renewable_daily_energy_available = total_renewable_available.resample("D").sum()

    # Alinear índices
    common_index = renewable_daily_energy_used.index.intersection(
        renewable_daily_energy_available.index
    )

    if len(common_index) == 0:
        return None

    renewable_daily_energy_used = renewable_daily_energy_used.loc[common_index]
    renewable_daily_energy_available = renewable_daily_energy_available.loc[common_index]

    fig, ax = plt.subplots(figsize=(12, 5))

    # Curva usada
    ax.step(
        renewable_daily_energy_used.index,
        renewable_daily_energy_used.values,
        where="mid",
        linewidth=1.8,
        label="Daily renewable energy (used)"
    )

    # Curva disponible
    ax.step(
        renewable_daily_energy_available.index,
        renewable_daily_energy_available.values,
        where="mid",
        linewidth=1.8,
        linestyle="--",
        label="Daily renewable energy (available)"
    )

    # Área usada
    ax.fill_between(
        renewable_daily_energy_used.index,
        0,
        renewable_daily_energy_used.values,
        step="mid",
        alpha=0.25
    )

    # (Opcional) área de curtailment
    ax.fill_between(
        renewable_daily_energy_used.index,
        renewable_daily_energy_used.values,
        renewable_daily_energy_available.values,
        step="mid",
        alpha=0.15,
        label="Curtailment"
    )

    ax.set_title("Daily renewable energy")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/day]")

    # Formato eje X adaptativo
    n_days = len(renewable_daily_energy_used)

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

def plot_total_renewable_weekly_energy(
    dispatch_clean: pd.DataFrame,
    df_available_renewable: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con dos curvas [MWh/week]:
    - Energía renovable generada semanal (PV + Wind)
    - Energía renovable disponible semanal
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = dispatch_clean["PV"] if "PV" in dispatch_clean.columns else 0
    wind = dispatch_clean["Wind"] if "Wind" in dispatch_clean.columns else 0

    # Potencia renovable usada
    total_renewable_used = pv + wind

    # Energía semanal usada
    renewable_weekly_energy_used = total_renewable_used.resample("W").sum()

    # Renovable disponible
    if df_available_renewable is None or df_available_renewable.empty:
        return None

    total_renewable_available = df_available_renewable.sum(axis=1)
    renewable_weekly_energy_available = total_renewable_available.resample("W").sum()

    # Alinear índices
    common_index = renewable_weekly_energy_used.index.intersection(
        renewable_weekly_energy_available.index
    )

    if len(common_index) == 0:
        return None

    renewable_weekly_energy_used = renewable_weekly_energy_used.loc[common_index]
    renewable_weekly_energy_available = renewable_weekly_energy_available.loc[common_index]

    fig, ax = plt.subplots(figsize=(12, 5))

    # Curva usada
    ax.step(
        renewable_weekly_energy_used.index,
        renewable_weekly_energy_used.values,
        where="mid",
        linewidth=1.8,
        label="Weekly renewable energy (used)"
    )

    # Curva disponible
    ax.step(
        renewable_weekly_energy_available.index,
        renewable_weekly_energy_available.values,
        where="mid",
        linewidth=1.8,
        linestyle="--",
        label="Weekly renewable energy (available)"
    )

    # Área usada
    ax.fill_between(
        renewable_weekly_energy_used.index,
        0,
        renewable_weekly_energy_used.values,
        step="mid",
        alpha=0.25
    )

    # Área curtailment (solo si tiene sentido)
    ax.fill_between(
        renewable_weekly_energy_used.index,
        renewable_weekly_energy_used.values,
        renewable_weekly_energy_available.values,
        step="mid",
        alpha=0.15,
        label="Curtailment"
    )

    ax.set_title("Weekly renewable generation")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/week]")

    # Formato eje X adaptativo
    n_weeks = len(renewable_weekly_energy_used)

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

def plot_pv_wind_power(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - potencia PV por hora [MW]
    - potencia Wind por hora [MW]
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = (
        dispatch_clean["PV"]
        if "PV" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    wind = (
        dispatch_clean["Wind"]
        if "Wind" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    fig, ax = plt.subplots(figsize=(12, 5))

    # Curva PV
    ax.plot(
        pv.index,
        pv.values,
        linewidth=1.8,
        label="PV"
    )

    # Curva Wind
    ax.plot(
        wind.index,
        wind.values,
        linewidth=1.8,
        label="Wind"
    )

    ax.set_title("Renewable generation by source")
    ax.set_xlabel("Time")
    ax.set_ylabel("Power [MW]")

    # Formato eje X adaptativo (igual que tu función)
    n_snapshots = len(pv)

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

def plot_pv_wind_daily_energy(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - energía diaria PV [MWh/day]
    - energía diaria Wind [MWh/day]
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = (
        dispatch_clean["PV"]
        if "PV" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    wind = (
        dispatch_clean["Wind"]
        if "Wind" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    # Energía diaria (asumiendo datos horarios → suma = MWh/día)
    pv_daily_energy = pv.resample("D").sum()
    wind_daily_energy = wind.resample("D").sum()

    fig, ax = plt.subplots(figsize=(12, 5))

    # PV
    ax.step(
        pv_daily_energy.index,
        pv_daily_energy.values,
        where="mid",
        linewidth=1.8,
        label="Daily PV energy"
    )
    ax.fill_between(
        pv_daily_energy.index,
        0,
        pv_daily_energy.values,
        step="mid",
        alpha=0.25
    )

    # Wind
    ax.step(
        wind_daily_energy.index,
        wind_daily_energy.values,
        where="mid",
        linewidth=1.8,
        label="Daily Wind energy"
    )
    ax.fill_between(
        wind_daily_energy.index,
        0,
        wind_daily_energy.values,
        step="mid",
        alpha=0.25
    )

    ax.set_title("Daily renewable generation by source")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/day]")

    # Formato eje X adaptativo
    n_days = len(pv_daily_energy)

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

def plot_pv_wind_weekly_energy(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - energía semanal PV [MWh/week]
    - energía semanal Wind [MWh/week]
    """

    if horizon != "Multiperiod":
        return None

    # Comprobación de columnas
    if "PV" not in dispatch_clean.columns and "Wind" not in dispatch_clean.columns:
        return None

    pv = (
        dispatch_clean["PV"]
        if "PV" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    wind = (
        dispatch_clean["Wind"]
        if "Wind" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    # Energía semanal
    pv_weekly_energy = pv.resample("W").sum()
    wind_weekly_energy = wind.resample("W").sum()

    fig, ax = plt.subplots(figsize=(12, 5))

    # PV
    ax.step(
        pv_weekly_energy.index,
        pv_weekly_energy.values,
        where="mid",
        linewidth=1.8,
        label="Weekly PV energy"
    )
    ax.fill_between(
        pv_weekly_energy.index,
        0,
        pv_weekly_energy.values,
        step="mid",
        alpha=0.25
    )

    # Wind
    ax.step(
        wind_weekly_energy.index,
        wind_weekly_energy.values,
        where="mid",
        linewidth=1.8,
        label="Weekly Wind energy"
    )
    ax.fill_between(
        wind_weekly_energy.index,
        0,
        wind_weekly_energy.values,
        step="mid",
        alpha=0.25
    )

    ax.set_title("Weekly renewable generation by source")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/week]")

    # Formato eje X adaptativo (igual que el tuyo)
    n_weeks = len(pv_weekly_energy)

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


def renewable_graph_resolution_choice(
    df_SYS_settings: pd.DataFrame,
    dispatch_clean: pd.DataFrame,
    df_available_renewable: pd.DataFrame,
) -> tuple[Figure | None, Figure | None]:
    
    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    resolution = df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]
    n_snapshots = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            fig_total = plot_total_renewable_weekly_energy(dispatch_clean, df_available_renewable, horizon)
            fig_pv_wind = plot_pv_wind_weekly_energy(dispatch_clean, horizon)
        elif 60 <= n_snapshots < 200:
            fig_total = plot_total_renewable_daily_energy(dispatch_clean, df_available_renewable, horizon)
            fig_pv_wind = plot_pv_wind_daily_energy(dispatch_clean, horizon)
        else:
            fig_total = plot_total_renewable_power(dispatch_clean, df_available_renewable, horizon)
            fig_pv_wind = plot_pv_wind_power(dispatch_clean, horizon)

    elif resolution == "Hourly":
        fig_total = plot_total_renewable_power(dispatch_clean, df_available_renewable, horizon)
        fig_pv_wind = plot_pv_wind_power(dispatch_clean, horizon)

    elif resolution == "Daily":
        fig_total = plot_total_renewable_daily_energy(dispatch_clean, df_available_renewable, horizon)
        fig_pv_wind = plot_pv_wind_daily_energy(dispatch_clean, horizon)

    elif resolution == "Weekly":
        fig_total = plot_total_renewable_weekly_energy(dispatch_clean, df_available_renewable, horizon)
        fig_pv_wind = plot_pv_wind_weekly_energy(dispatch_clean, horizon)

    else:
        fig_total = None
        fig_pv_wind = None

    return fig_total, fig_pv_wind


