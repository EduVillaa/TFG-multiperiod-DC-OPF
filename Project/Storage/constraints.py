import pypsa

def add_battery_constraints(grid: pypsa.Network, snapshots, battery_specs: list[dict]) -> None:

    m = grid.model

    # Si ninguna batería es optimizable, salir sin hacer nada
    if not any(b["optimize_battery"] for b in battery_specs):
        return

    # Seguridad extra: comprobar que las variables existen
    if "Store-e_nom" not in m.variables:
        return
    if "Link-p_nom" not in m.variables:
        return

    store_e_nom = m.variables["Store-e_nom"]
    link_p_nom = m.variables["Link-p_nom"]

    for bat in battery_specs:
        if not bat["optimize_battery"]:
            continue

        store_name = bat["store_name"]
        charge_link_name = bat["charge_link_name"]
        discharge_link_name = bat["discharge_link_name"]
        max_hours = bat["max_hours"]

        m.add_constraints(
            store_e_nom.loc[store_name] == max_hours * link_p_nom.loc[discharge_link_name],
            name=f"battery_duration_{store_name}"
        )

        m.add_constraints(
            link_p_nom.loc[charge_link_name] == link_p_nom.loc[discharge_link_name],
            name=f"battery_symmetric_power_{store_name}"
        )
