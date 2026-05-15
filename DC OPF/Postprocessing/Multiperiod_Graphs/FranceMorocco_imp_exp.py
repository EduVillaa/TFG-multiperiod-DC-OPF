import pandas as pd
import pypsa
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


def build_interconnection_clean(grid: pypsa.Network) -> pd.DataFrame:
    """
    Construye un DataFrame con los intercambios con Francia y Marruecos
    a partir de los Links LPCC_*.

    Convención asumida:
        bus0 = PCC externo
        bus1 = bus español/portugués

    Entonces:
        p0 > 0  -> importación desde el país externo
        p0 < 0  -> exportación hacia el país externo

    Devuelve columnas:
        France_import   (positiva)
        France_export   (negativa)
        Morocco_import  (positiva)
        Morocco_export  (negativa)
    """

    df = pd.DataFrame(index=grid.snapshots)

    country_prefix = {
        "France": "LPCC_France_",
        "Morocco": "LPCC_Morocco_",
    }

    for country, prefix in country_prefix.items():
        cols = [c for c in grid.links_t.p0.columns if c.startswith(prefix)]

        if cols:
            flow = grid.links_t.p0[cols].sum(axis=1)
            df[f"{country}_import"] = grid.links_t.p0[cols].clip(lower=0).sum(axis=1)
            df[f"{country}_export"] = grid.links_t.p0[cols].clip(upper=0).sum(axis=1)
        else:
            df[f"{country}_import"] = 0.0
            df[f"{country}_export"] = 0.0

    # Elimina columnas completamente nulas
    df = df.loc[:, (df.abs() > 1e-6).any()]

    return df

def plot_interconnection_figure_hourly_snapshots(
    interconnection_clean: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve una figura con los intercambios horarios:
    - importaciones positivas
    - exportaciones negativas
    - Francia en azul
    - Marruecos en rojo
    """

    if horizon != "Multiperiod":
        return None

    fig, ax = plt.subplots(figsize=(14, 6))

    countries = ["France", "Morocco"]
    colors = {
        "France": "#42A5F5",   # azul
        "Morocco": "#EF5350",  # rojo
    }

    # ----------------------------
    # Positivos: importaciones
    # ----------------------------
    base_pos = pd.Series(0.0, index=interconnection_clean.index)

    for country in countries:
        col = f"{country}_import"
        if col in interconnection_clean.columns:
            y = interconnection_clean[col]

            ax.fill_between(
                interconnection_clean.index,
                base_pos,
                base_pos + y,
                step="post",
                alpha=0.85,
                label=country,
                color=colors[country]
            )
            ax.step(
                interconnection_clean.index,
                base_pos + y,
                where="post",
                color=colors[country],
                linewidth=1
            )
            base_pos += y

    # ----------------------------
    # Negativos: exportaciones
    # ----------------------------
    base_neg = pd.Series(0.0, index=interconnection_clean.index)

    for country in countries:
        col = f"{country}_export"
        if col in interconnection_clean.columns:
            y = interconnection_clean[col]   # ya es negativa

            ax.fill_between(
                interconnection_clean.index,
                base_neg,
                base_neg + y,
                step="post",
                alpha=0.85,
                label="_nolegend_",
                color=colors[country]
            )
            ax.step(
                interconnection_clean.index,
                base_neg + y,
                where="post",
                color=colors[country],
                linewidth=1
            )
            base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Interconnection exchange")
    ax.set_ylabel("Power [MW]")
    ax.set_xlabel("Time")

    # Formato eje X
    n_snapshots = len(interconnection_clean)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        interval = max(1, int(1 / 15 * n_snapshots - 1 / 5))
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%Hh\n%d %b"))

    else:
        interval = max(1, int(1 / 408 * n_snapshots + 9 / 17))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_interconnection_figure_daily_average(
    interconnection_clean: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve la figura del intercambio diario agregado
    en MWh/día.
    """

    if horizon != "Multiperiod":
        return None

    interconnection_daily = interconnection_clean.resample("D").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    countries = ["France", "Morocco"]
    colors = {
        "France": "#42A5F5",
        "Morocco": "#EF5350",
    }

    # Positivos
    base_pos = pd.Series(0.0, index=interconnection_daily.index)
    for country in countries:
        col = f"{country}_import"
        if col in interconnection_daily.columns:
            y = interconnection_daily[col]
            ax.fill_between(
                interconnection_daily.index,
                base_pos,
                base_pos + y,
                step="mid",
                alpha=0.85,
                label=country,
                color=colors[country]
            )
            ax.step(
                interconnection_daily.index,
                base_pos + y,
                where="mid",
                color=colors[country],
                linewidth=1
            )
            base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=interconnection_daily.index)
    for country in countries:
        col = f"{country}_export"
        if col in interconnection_daily.columns:
            y = interconnection_daily[col]
            ax.fill_between(
                interconnection_daily.index,
                base_neg,
                base_neg + y,
                step="mid",
                alpha=0.85,
                label="_nolegend_",
                color=colors[country]
            )
            ax.step(
                interconnection_daily.index,
                base_neg + y,
                where="mid",
                color=colors[country],
                linewidth=1
            )
            base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Daily interconnection exchange")
    ax.set_ylabel("Energy [MWh/day]")
    ax.set_xlabel("Time")

    n_days = len(interconnection_daily)
    ax.xaxis.set_major_locator(mdates.DayLocator(interval=max(1, int(7 / 130 * n_days))))
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d\n%b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def plot_interconnection_figure_weekly_average(
    interconnection_clean: pd.DataFrame,
    horizon: str = "Multiperiod"
):
    """
    Devuelve la figura del intercambio semanal agregado
    en MWh/semana.
    """

    if horizon != "Multiperiod":
        return None

    interconnection_week = interconnection_clean.resample("W").sum()

    fig, ax = plt.subplots(figsize=(14, 6))

    countries = ["France", "Morocco"]
    colors = {
        "France": "#42A5F5",
        "Morocco": "#EF5350",
    }

    # Positivos
    base_pos = pd.Series(0.0, index=interconnection_week.index)
    for country in countries:
        col = f"{country}_import"
        if col in interconnection_week.columns:
            y = interconnection_week[col]
            ax.fill_between(
                interconnection_week.index,
                base_pos,
                base_pos + y,
                step="mid",
                alpha=0.85,
                label=country,
                color=colors[country]
            )
            ax.step(
                interconnection_week.index,
                base_pos + y,
                where="mid",
                color=colors[country],
                linewidth=1
            )
            base_pos += y

    # Negativos
    base_neg = pd.Series(0.0, index=interconnection_week.index)
    for country in countries:
        col = f"{country}_export"
        if col in interconnection_week.columns:
            y = interconnection_week[col]
            ax.fill_between(
                interconnection_week.index,
                base_neg,
                base_neg + y,
                step="mid",
                alpha=0.85,
                label="_nolegend_",
                color=colors[country]
            )
            ax.step(
                interconnection_week.index,
                base_neg + y,
                where="mid",
                color=colors[country],
                linewidth=1
            )
            base_neg += y

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Weekly interconnection exchange")
    ax.set_ylabel("Energy [MWh/week]")
    ax.set_xlabel("Time")

    n_weeks = len(interconnection_week)
    ax.xaxis.set_major_locator(
        mdates.WeekdayLocator(interval=int(13 / 127 * n_weeks - 136 / 127) + 1)
    )
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    fig.tight_layout()

    return fig

def interconnection_graph_resolution_choice(
    df_SYS_settings: pd.DataFrame,
    grid: pypsa.Network
):
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    horizon = params["Static / Multiperiod"]
    resolution = params["Graph resolution"]
    n_days = params["Simulation duration (days)"]

    interconnection_clean = build_interconnection_clean(grid)

    if resolution == "Auto":
        if n_days >= 200:
            return plot_interconnection_figure_weekly_average(interconnection_clean, horizon)
        elif 60 <= n_days < 200:
            return plot_interconnection_figure_daily_average(interconnection_clean, horizon)
        else:
            return plot_interconnection_figure_hourly_snapshots(interconnection_clean, horizon)

    elif resolution == "Hourly":
        return plot_interconnection_figure_hourly_snapshots(interconnection_clean, horizon)

    elif resolution == "Daily":
        return plot_interconnection_figure_daily_average(interconnection_clean, horizon)

    elif resolution == "Weekly":
        return plot_interconnection_figure_weekly_average(interconnection_clean, horizon)

    return None