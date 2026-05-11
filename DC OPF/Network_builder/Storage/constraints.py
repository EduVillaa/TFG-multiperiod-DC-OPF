import pypsa



# Esta función se asegura de que cuando la potencia del inversor es extendable, los links de carga y descarga al optimizarse
# tengan los dos la misma potencia
def add_battery_constraints(grid: pypsa.Network, snapshots, battery_specs: list[dict]) -> None:
    m = grid.model

    if not any(b["optimize_p"] for b in battery_specs):
        return

    if "Link-p_nom" not in m.variables:
        return

    link_p_nom = m.variables["Link-p_nom"]

    for bat in battery_specs:
        if not bat["optimize_p"]:
            continue

        store_name = bat["store_name"]
        charge_link_name = bat["charge_link_name"]
        discharge_link_name = bat["discharge_link_name"]

        m.add_constraints(
            link_p_nom.loc[charge_link_name] == link_p_nom.loc[discharge_link_name],
            name=f"battery_symmetric_power_{store_name}"
        )