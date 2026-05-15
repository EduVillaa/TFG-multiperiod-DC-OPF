import plotly.graph_objects as go
import pandas as pd


def plot_generation_mix_pie(
    dispatch_clean,
    include_storage_discharge=True,
    include_grid_import=True,
    include_shedding=False,
    min_percentage=0.5
):
    """
    Genera un gráfico tipo queso con el mix de generación acumulado
    durante el horizonte de simulación.

    Parameters
    ----------
    dispatch_clean : pd.DataFrame
        DataFrame con las series temporales de generación/despacho.
        Se espera que tenga columnas como:
        PV, Wind, ror, Hidroelectric_discharge, biomass, Nuclear,
        CCGT, Other, Grid_import, battery_discharge, shedding.

    include_storage_discharge : bool
        Si True, incluye la descarga de baterías en el mix.

    include_grid_import : bool
        Si True, incluye las importaciones de red en el mix.

    include_shedding : bool
        Si True, incluye la energía no suministrada como categoría separada.

    min_percentage : float
        Porcentaje mínimo para mostrar una categoría individualmente.
        Las categorías por debajo de este porcentaje se agrupan en "Other small sources".

    Returns
    -------
    fig : plotly.graph_objects.Figure
        Figura Plotly del mix de generación.

    mix_df : pd.DataFrame
        Tabla con energía en MWh y porcentaje por tecnología.
    """

    def safe_sum(df, col):
        return df[col].sum() if col in df.columns else 0.0

    # --------------------------------------------------
    # Energía generada / aportada durante el horizonte
    # --------------------------------------------------
    generation_data = {
        "PV": safe_sum(dispatch_clean, "PV"),
        "Wind": safe_sum(dispatch_clean, "Wind"),
        "Run-of-river hydro": safe_sum(dispatch_clean, "ror"),
        "Reservoir hydro discharge": safe_sum(dispatch_clean, "Hidroelectric_discharge"),
        "Biomass": safe_sum(dispatch_clean, "biomass"),
        "Nuclear": safe_sum(dispatch_clean, "Nuclear"),
        "CCGT": safe_sum(dispatch_clean, "CCGT"),
        "Other": safe_sum(dispatch_clean, "Other"),
    }

    if include_storage_discharge:
        generation_data["Battery discharge"] = safe_sum(dispatch_clean, "battery_discharge")

    if include_grid_import:
        generation_data["Grid import"] = safe_sum(dispatch_clean, "Grid_import")

    if include_shedding:
        generation_data["Unserved load"] = safe_sum(dispatch_clean, "shedding")

    # Eliminar valores nulos, negativos o numéricamente despreciables
    generation_data = {
        k: max(v, 0.0)
        for k, v in generation_data.items()
        if v > 1e-9
    }

    total_generation = sum(generation_data.values())

    if total_generation <= 1e-9:
        raise ValueError("No hay generación positiva para representar en el mix.")

    # --------------------------------------------------
    # Crear DataFrame resumen
    # --------------------------------------------------
    mix_df = pd.DataFrame({
        "Technology": list(generation_data.keys()),
        "Energy (MWh)": list(generation_data.values())
    })

    mix_df["Share (%)"] = 100 * mix_df["Energy (MWh)"] / total_generation

    mix_df = mix_df.sort_values("Energy (MWh)", ascending=False).reset_index(drop=True)

    # --------------------------------------------------
    # Agrupar tecnologías pequeñas
    # --------------------------------------------------
    large_sources = mix_df[mix_df["Share (%)"] >= min_percentage].copy()
    small_sources = mix_df[mix_df["Share (%)"] < min_percentage].copy()

    if not small_sources.empty:
        small_row = pd.DataFrame({
            "Technology": ["Other small sources"],
            "Energy (MWh)": [small_sources["Energy (MWh)"].sum()],
            "Share (%)": [small_sources["Share (%)"].sum()]
        })

        plot_df = pd.concat([large_sources, small_row], ignore_index=True)
    else:
        plot_df = large_sources.copy()

    # --------------------------------------------------
    # Colores
    # --------------------------------------------------
    color_map = {
        "PV": "#FFD54F",
        "Wind": "#4FC3F7",
        "Run-of-river hydro": "#4DB6AC",
        "Reservoir hydro discharge": "#1976D2",
        "Battery discharge": "#66BB6A",
        "Biomass": "#8D6E63",
        "Nuclear": "#9575CD",
        "CCGT": "#E57373",
        "Other": "#EF5350",
        "Grid import": "#B0BEC5",
        "Unserved load": "#D50000",
        "Other small sources": "#9E9E9E",
    }

    colors = [
        color_map.get(tech, "#BDBDBD")
        for tech in plot_df["Technology"]
    ]

    # --------------------------------------------------
    # Gráfico tipo queso
    # --------------------------------------------------
    fig = go.Figure(data=[
        go.Pie(
            labels=plot_df["Technology"],
            values=plot_df["Energy (MWh)"],
            hole=0.35,
            marker=dict(
                colors=colors,
                line=dict(color="white", width=2)
            ),
            textinfo="label+percent",
            hovertemplate=(
                "<b>%{label}</b><br>"
                "Energy: %{value:,.2f} MWh<br>"
                "Share: %{percent}<extra></extra>"
            )
        )
    ])

    fig.update_layout(
        title=dict(
            text="Generation mix over the simulation horizon",
            x=0.03,
            y=0.95
        ),
        font=dict(size=16),
        height=650,
        margin=dict(l=20, r=20, t=90, b=20),
        legend=dict(
            orientation="v",
            yanchor="middle",
            y=0.5,
            xanchor="left",
            x=1.02
        )
    )

    # Redondear tabla final
    mix_df["Energy (MWh)"] = mix_df["Energy (MWh)"].round(2)
    mix_df["Share (%)"] = mix_df["Share (%)"].round(2)

    return fig, mix_df