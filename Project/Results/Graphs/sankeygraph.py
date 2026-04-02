import plotly.graph_objects as go
import pandas as pd



def plot_energy_balance_sankey(dispatch_clean, grid, df_available_renewable):


    def safe_sum(df, col):
        return df[col].sum() if col in df.columns else 0.0

    pv = safe_sum(dispatch_clean, "PV")
    wind = safe_sum(dispatch_clean, "Wind")
    grid_import = safe_sum(dispatch_clean, "Grid_import")
    battery_discharge = safe_sum(dispatch_clean, "battery_discharge")
    dispatch = safe_sum(dispatch_clean, "Dispatch")
    shedding = safe_sum(dispatch_clean, "shedding")

    load = grid.loads_t.p.sum().sum() if hasattr(grid.loads_t.p, "sum") else 0.0
    battery_charge = -safe_sum(dispatch_clean, "battery_charge")
    grid_export = -safe_sum(dispatch_clean, "Grid_export")
    battery_loses = battery_charge-battery_discharge


    total_in = pv + wind + grid_import + battery_discharge + dispatch + shedding
    total_out = load + battery_charge + grid_export

    if battery_charge > 1e-9:
        percentage_losses_battery = round(battery_loses / battery_charge * 100, 2)
        percentage_losses_battery_str = f"{percentage_losses_battery} %"
    else:
        percentage_losses_battery_str = "N/A"

    labels = [
        "PV",                 
        "Wind",               
        "Grid import",        
        "Battery discharge",  
        "Dispatch",          
        "Energy supplied",    
        "Served load",               
        "Battery charge",     
        "Grid export",        
    ]

    source = [
        0, 1, 2, 3, 4,
        5, 5, 5
    ]

    target = [
        5, 5, 5, 5, 5,
        6, 7, 8
    ]

    value = [
        pv, wind, grid_import, battery_discharge, dispatch,
        load-shedding, battery_charge, grid_export
    ]

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=25,
            thickness=20,
            line=dict(color="black", width=0.5),
            label=labels
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

    available_total = df_available_renewable.sum().sum()

    sankey_df = pd.DataFrame({
    "Category": [
        "",
        "",
        "PV",
        "Wind",
        "Dispatch",
        "Grid import",
        "Battery discharge",
        "Total",
        "",
        "",
        "Served load",
        "Unserved load",
        "Grid export",
        "Battery charge",
        "Total",
        "",
        "",
        "",
        "",
        "",
        "",
        "Renewable share",
        "Total available renewable (MWh)",
        "Total used renewable (MWh)",
        "Curtailment (MWh)",
        "Curtailment (%)"
    ],
    "Value (MWh)": [
        "",
        "Total generation (MWh)",
        round(pv, 2),
        round(wind, 2),
        round(dispatch, 2),
        round(grid_import, 2),
        round(battery_discharge, 2),
        round(pv + wind + dispatch + grid_import + battery_discharge, 2),
        None,
        "Total consumption (MWh)",
        round(load-shedding, 2),
        round(shedding, 2),
        round(grid_export, 2),
        round(battery_charge, 2),
        round(load + grid_export + battery_charge - shedding, 2),
        "",
        "Total battery loses (MWh)",
        round(battery_loses, 2),
        percentage_losses_battery_str,
        "",
        "Renewable data",
        f"{round(100 * (pv + wind) / load, 2)} %",
        round(available_total, 2),
        round(pv+wind, 2),
        round(available_total-pv-wind, 2),
        round((available_total-pv-wind)/available_total, 2)

    ]
    })

    #fig.show()
    return fig, sankey_df

