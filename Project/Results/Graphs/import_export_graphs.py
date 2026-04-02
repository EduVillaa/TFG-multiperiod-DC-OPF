import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.figure import Figure


def GridExportImport_per_hour_graph(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - importación de la red
    - exportación a la red
    - intercambio neto con la red

    Se asume que:
    - Grid_import >= 0
    - Grid_export <= 0
    """

    if horizon != "Multiperiod":
        return None

    if "Grid_import" not in dispatch_clean.columns and "Grid_export" not in dispatch_clean.columns:
        return None

    fig, ax = plt.subplots(figsize=(12, 5))

    grid_import = dispatch_clean["Grid_import"] if "Grid_import" in dispatch_clean.columns else pd.Series(0.0, index=dispatch_clean.index)
    grid_export = dispatch_clean["Grid_export"] if "Grid_export" in dispatch_clean.columns else pd.Series(0.0, index=dispatch_clean.index)

    # Intercambio neto con la red
    net_grid_exchange = grid_import + grid_export

    # Importación
    if "Grid_import" in dispatch_clean.columns:
        ax.step(
            grid_import.index,
            grid_import.values,
            where="post",
            label="Grid import",
            linewidth=1.8
        )
        ax.fill_between(
            grid_import.index,
            0,
            grid_import.values,
            step="post",
            alpha=0.25
        )

    # Exportación
    if "Grid_export" in dispatch_clean.columns:
        ax.step(
            grid_export.index,
            grid_export.values,
            where="post",
            label="Grid export",
            linewidth=1.8
        )
        ax.fill_between(
            grid_export.index,
            0,
            grid_export.values,
            step="post",
            alpha=0.25
        )

    # Intercambio neto
    ax.step(
        net_grid_exchange.index,
        net_grid_exchange.values,
        where="post",
        linestyle="--",
        linewidth=2,
        label="Net grid exchange"
    )

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Grid import/export over time")
    ax.set_xlabel("Time")
    ax.set_ylabel("Power [MW]")

    n_snapshots = len(dispatch_clean)

    if n_snapshots <= 24:
        ax.xaxis.set_major_locator(mdates.HourLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%H:%M"))

    elif n_snapshots <= 24 * 7:
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))
        ax.xaxis.set_minor_locator(mdates.HourLocator(interval=12))
        ax.xaxis.set_minor_formatter(mdates.DateFormatter("%Hh"))
        ax.tick_params(axis="x", which="major", pad=15)
        ax.tick_params(axis="x", which="minor", pad=3)

    else:
        interval = max(1, int(n_snapshots / 24 / 14))
        ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def GridExportImport_daily_energy_graph(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - energía total importada cada día [MWh/day]
    - energía total exportada cada día [MWh/day]
    - energía neta intercambiada cada día [MWh/day]

    Se asume que:
    - Grid_import >= 0
    - Grid_export <= 0
    - los snapshots son horarios
    """

    if horizon != "Multiperiod":
        return None

    if "Grid_import" not in dispatch_clean.columns and "Grid_export" not in dispatch_clean.columns:
        return None

    grid_import = (
        dispatch_clean["Grid_import"]
        if "Grid_import" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    grid_export = (
        dispatch_clean["Grid_export"]
        if "Grid_export" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    # Energía diaria
    import_daily_energy = grid_import.resample("D").sum()
    export_daily_energy = grid_export.resample("D").sum()   # negativa para representar exportación
    net_daily_energy = import_daily_energy + export_daily_energy

    fig, ax = plt.subplots(figsize=(12, 5))

    # Importación diaria
    ax.step(
        import_daily_energy.index,
        import_daily_energy.values,
        where="mid",
        linewidth=1.8,
        label="Daily imported energy"
    )
    ax.fill_between(
        import_daily_energy.index,
        0,
        import_daily_energy.values,
        step="mid",
        alpha=0.25
    )

    # Exportación diaria
    ax.step(
        export_daily_energy.index,
        export_daily_energy.values,
        where="mid",
        linewidth=1.8,
        label="Daily exported energy"
    )
    ax.fill_between(
        export_daily_energy.index,
        0,
        export_daily_energy.values,
        step="mid",
        alpha=0.25
    )

    # Neto diario
    ax.step(
        net_daily_energy.index,
        net_daily_energy.values,
        where="mid",
        linestyle="--",
        linewidth=2.2,
        label="Daily net exchanged energy"
    )

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Daily grid exchanged energy")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/day]")

    n_days = len(import_daily_energy)

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

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig

def GridExportImport_weekly_energy_graph(dispatch_clean: pd.DataFrame, horizon: str = "Multiperiod"):
    """
    Devuelve una figura con:
    - energía total importada cada semana [MWh/week]
    - energía total exportada cada semana [MWh/week]
    - energía neta intercambiada cada semana [MWh/week]

    Se asume que:
    - Grid_import >= 0
    - Grid_export <= 0
    - los snapshots son horarios
    """

    if horizon != "Multiperiod":
        return None

    if "Grid_import" not in dispatch_clean.columns and "Grid_export" not in dispatch_clean.columns:
        return None

    grid_import = (
        dispatch_clean["Grid_import"]
        if "Grid_import" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    grid_export = (
        dispatch_clean["Grid_export"]
        if "Grid_export" in dispatch_clean.columns
        else pd.Series(0.0, index=dispatch_clean.index)
    )

    # Energía semanal
    import_weekly_energy = grid_import.resample("W").sum()
    export_weekly_energy = grid_export.resample("W").sum()   # negativa para representar exportación
    net_weekly_energy = import_weekly_energy + export_weekly_energy

    fig, ax = plt.subplots(figsize=(12, 5))

    # Importación semanal
    ax.step(
        import_weekly_energy.index,
        import_weekly_energy.values,
        where="mid",
        linewidth=1.8,
        label="Weekly imported energy"
    )
    ax.fill_between(
        import_weekly_energy.index,
        0,
        import_weekly_energy.values,
        step="mid",
        alpha=0.25
    )

    # Exportación semanal
    ax.step(
        export_weekly_energy.index,
        export_weekly_energy.values,
        where="mid",
        linewidth=1.8,
        label="Weekly exported energy"
    )
    ax.fill_between(
        export_weekly_energy.index,
        0,
        export_weekly_energy.values,
        step="mid",
        alpha=0.25
    )

    # Neto semanal
    ax.step(
        net_weekly_energy.index,
        net_weekly_energy.values,
        where="mid",
        linestyle="--",
        linewidth=2.2,
        label="Weekly net exchanged energy"
    )

    ax.axhline(0, color="black", linewidth=1)
    ax.set_title("Weekly grid exchanged energy")
    ax.set_xlabel("Time")
    ax.set_ylabel("Energy [MWh/week]")

    n_weeks = len(import_weekly_energy)

    if n_weeks <= 12:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=1))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b"))

    elif n_weeks <= 52:
        ax.xaxis.set_major_locator(mdates.WeekdayLocator(interval=4))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%d %b\n%Y"))

    else:
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=2))
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b\n%Y"))

    ax.legend(loc="upper left", bbox_to_anchor=(1.02, 1))
    ax.grid(True, axis="y")
    fig.tight_layout()

    return fig


def GridExportImport_graph_resolution_choice(
    df_SYS_settings, 
    dispatch_clean: pd.DataFrame
) -> Figure | None:

    horizon = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]
    resolution = df_SYS_settings.loc[7, "SYSTEM PARAMETERS"]
    n_snapshots = df_SYS_settings.loc[6, "SYSTEM PARAMETERS"]

    if resolution == "Auto":
        if n_snapshots >= 200:
            return GridExportImport_weekly_energy_graph(dispatch_clean, horizon)
        elif 60 <= n_snapshots < 200:
            return GridExportImport_daily_energy_graph(dispatch_clean, horizon)
        else:
            return GridExportImport_per_hour_graph(dispatch_clean, horizon)

    elif resolution == "Hourly":
        return GridExportImport_per_hour_graph(dispatch_clean, horizon)

    elif resolution == "Daily":
        return GridExportImport_daily_energy_graph(dispatch_clean, horizon)

    elif resolution == "Weekly":
        return GridExportImport_weekly_energy_graph(dispatch_clean, horizon)

    return None