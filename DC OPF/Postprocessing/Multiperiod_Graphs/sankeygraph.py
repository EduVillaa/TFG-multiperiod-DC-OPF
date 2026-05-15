import plotly.graph_objects as go
import pandas as pd


def plot_energy_balance_sankey(dispatch_clean, grid, df_available_renewable, CFsolar, CFwind):

    def safe_sum(df, col):
        return df[col].sum() if col in df.columns else 0.0

    # -----------------------------
    # Energía generada / aportada
    # -----------------------------
    pv = safe_sum(dispatch_clean, "PV")
    wind = safe_sum(dispatch_clean, "Wind")
    ror = safe_sum(dispatch_clean, "ror")
    biomass = safe_sum(dispatch_clean, "biomass")
    nuclear = safe_sum(dispatch_clean, "Nuclear")
    ccgt = safe_sum(dispatch_clean, "CCGT")
    other = safe_sum(dispatch_clean, "Other")
    grid_import = safe_sum(dispatch_clean, "Grid_import")
    battery_discharge = safe_sum(dispatch_clean, "battery_discharge")
    hydro_discharge = safe_sum(dispatch_clean, "Hidroelectric_discharge")
    shedding = safe_sum(dispatch_clean, "shedding")

    # -----------------------------
    # Consumos / salidas
    # -----------------------------
    load = grid.loads_t.p.sum().sum() if hasattr(grid.loads_t.p, "sum") else 0.0

    battery_charge = -safe_sum(dispatch_clean, "battery_charge")
    grid_export = -safe_sum(dispatch_clean, "Grid_export")
    hydro_charge = -safe_sum(dispatch_clean, "Hidroelectric_charge")

    # Evitar valores negativos por posibles -0.0 o errores numéricos pequeños
    battery_charge = max(battery_charge, 0.0)
    grid_export = max(grid_export, 0.0)
    hydro_charge = max(hydro_charge, 0.0)

    battery_losses = battery_charge - battery_discharge
   

    # Si por redondeos salen pérdidas negativas pequeñas, las anulamos
    if abs(battery_losses) < 1e-6:
        battery_losses = 0.0

    # -----------------------------
    # Porcentajes de pérdidas
    # -----------------------------
    if battery_charge > 1e-9:
        percentage_losses_battery = round(battery_losses / battery_charge * 100, 2)
        percentage_losses_battery_str = f"{percentage_losses_battery} %"
    else:
        percentage_losses_battery_str = "N/A"


    # -----------------------------
    # Sankey
    # -----------------------------
    labels = [
        "PV",                       # 0
        "Wind",                     # 1
        "Run-of-river hydro",       # 2
        "Reservoir hydro discharge",# 3
        "Battery discharge",        # 4
        "Biomass",                  # 5
        "Nuclear",                  # 6
        "CCGT",                     # 7
        "Other",                    # 8
        "Grid import",              # 9
        "Energy supplied",          # 10
        "Served load",              # 11
        "Battery charge",           # 12
        "Hydroelectric charge",     # 13
        "Grid export",              # 14
        "Unserved load"             # 15
    ]

    raw_links = [
        # Sources to energy supplied
        (0, 10, pv),
        (1, 10, wind),
        (2, 10, ror),
        (3, 10, hydro_discharge),
        (4, 10, battery_discharge),
        (5, 10, biomass),
        (6, 10, nuclear),
        (7, 10, ccgt),
        (8, 10, other),
        (9, 10, grid_import),

        # Energy supplied to sinks
        (10, 11, load - shedding),
        (10, 12, battery_charge),
        (10, 13, hydro_charge),
        (10, 14, grid_export),
        (10, 15, shedding),
    ]

    # Eliminar flujos nulos o negativos por seguridad
    links = [(s, t, v) for s, t, v in raw_links if v > 1e-9]

    source = [s for s, t, v in links]
    target = [t for s, t, v in links]
    value = [v for s, t, v in links]

    node_colors = [
        "#FFD54F",  # PV
        "#4FC3F7",  # Wind
        "#4DB6AC",  # ror
        "#1976D2",  # hydro discharge
        "#66BB6A",  # battery discharge
        "#8D6E63",  # biomass
        "#9575CD",  # nuclear
        "#E57373",  # CCGT
        "#EF5350",  # other
        "#B0BEC5",  # import
        "#90A4AE",  # energy supplied
        "#C8E6C9",  # served load
        "#2E7D32",  # battery charge
        "#0D47A1",  # hydro charge
        "#424242",  # export
        "#D50000",  # unserved load
    ]

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=25,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=labels,
            color=node_colors
        ),
        link=dict(
            source=source,
            target=target,
            value=value
        )
    )])

    fig.update_layout(
        title=dict(
            text="Energy balance",
            x=0.03,
            y=0.95
        ),
        font=dict(size=16),
        height=750,
        margin=dict(l=20, r=20, t=100, b=20)
    )

    # -----------------------------
    # Indicadores
    # -----------------------------
    available_total = df_available_renewable.sum().sum()

    used_renewable = pv + wind + ror + biomass
    renewable_share = 100 * used_renewable / load if load > 0 else 0.0

    curtailment = available_total - pv - wind
    curtailment_percentage = 100 * curtailment / available_total if available_total > 0 else 0.0

    served_load = load - shedding
    served_load_percentage = 100 * served_load / load if load > 0 else 0.0

    total_generation = (
        pv
        + wind
        + ror
        + hydro_discharge
        + biomass
        + nuclear
        + ccgt
        + other
        + grid_import
        + battery_discharge
    )

    total_consumption = (
        served_load
        + shedding
        + battery_charge
        + hydro_charge
        + grid_export
    )

    sankey_df = pd.DataFrame({
        "Category": [
            "",
            "Total generation / supply (MWh)",
            "PV",
            "Wind",
            "Run-of-river hydro",
            "Reservoir hydro discharge",
            "Battery discharge",
            "Biomass",
            "Nuclear",
            "CCGT",
            "Other",
            "Grid import",
            "Total supply",
            "",
            "Total consumption / sinks (MWh)",
            "Served load",
            "Unserved load",
            "% of served load",
            "Battery charge",
            "Hydroelectric charge",
            "Grid export",
            "Total consumption",
            "",
            "Storage losses",
            "Total battery losses (MWh)",
            "Battery losses (%)",
            "",
            "Renewable data",
            "Renewable share",
            "Total available renewable (MWh)",
            "Total used renewable (MWh)",
            "Solar and wind curtailment (MWh)",
            "Solar and wind curtailment (%)",
            "Solar capacity factor",
            "Wind capacity factor",
        ],
        "Value": [
            "",
            "",
            round(pv, 2),
            round(wind, 2),
            round(ror, 2),
            round(hydro_discharge, 2),
            round(battery_discharge, 2),
            round(biomass, 2),
            round(nuclear, 2),
            round(ccgt, 2),
            round(other, 2),
            round(grid_import, 2),
            round(total_generation, 2),
            "",
            "",
            round(served_load, 2),
            round(shedding, 2),
            f"{round(served_load_percentage, 2)} %",
            round(battery_charge, 2),
            round(hydro_charge, 2),
            round(grid_export, 2),
            round(total_consumption, 2),
            "",
            "",
            round(battery_losses, 2),
            percentage_losses_battery_str,
            "",
            "",
            f"{round(renewable_share, 2)} %",
            round(available_total, 2),
            round(used_renewable, 2),
            round(curtailment, 2),
            f"{round(curtailment_percentage, 2)} %",
            round(CFsolar, 2),
            round(CFwind, 2),
        ]
    })

    return fig, sankey_df