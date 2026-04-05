import pypsa
import pandas as pd

def add_storage_as_store_links(
    grid: pypsa.Network,
    df_StorageUnit: pd.DataFrame
) -> list[dict]:
    """
    Añade baterías como Store + Link a la red.

    Si 'Optimize battery (0/1)' = 0:
        - batería fija

    Si 'Optimize battery (0/1)' = 1:
        - e_nom del Store extendable
        - p_nom de los Links extendable
        - las restricciones de acoplamiento se añaden luego con
          add_battery_constraints(...)

    Si 'Optimize battery (0/1)' = 1 y falta CAPEX de storage o inverter:
        - se toma 0.1 para ambos CAPEX
        - se usan límites máximos:
            p_nom_max = p_nom_base
            e_nom_max = e_nom_base
        - se lanza un único aviso:
            "Battery CAPEX missing. Using base values as maximum bounds."

    Devuelve:
        battery_specs: lista con metadatos de cada batería, necesaria
        para añadir restricciones extra al optimizador.
    """

    df = df_StorageUnit.copy()
    battery_specs = []
    missing_capex_warning_shown = False

    # -----------------------------
    # Conversión y valores por defecto
    # -----------------------------
    df["efficiency_store (p.u)"] = pd.to_numeric(
        df["efficiency_store (p.u)"], errors="coerce"
    ).fillna(0.95).astype(float)

    df["efficiency_dispatch (p.u)"] = pd.to_numeric(
        df["efficiency_dispatch (p.u)"], errors="coerce"
    ).fillna(0.95).astype(float)

    df["standing_loss (%/h)"] = pd.to_numeric(
        df["standing_loss (%/h)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["cyclic SOC (0/1)"] = pd.to_numeric(
        df["cyclic SOC (0/1)"], errors="coerce"
    ).fillna(1).astype(int)
    """
    df["initial SOC (%)"] = pd.to_numeric(
        df["initial SOC (%)"], errors="coerce"
    ).fillna(0.5).astype(float)
    """
    df["marginal_cost (€/MWh)"] = pd.to_numeric(
        df["marginal_cost (€/MWh)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["Rated active power (MW)"] = pd.to_numeric(
        df["Rated active power (MW)"], errors="coerce"
    )

    df["Max hours at rated active power (h)"] = pd.to_numeric(
        df["Max hours at rated active power (h)"], errors="coerce"
    )

    df["Optimize battery (0/1)"] = pd.to_numeric(
        df["Optimize battery (0/1)"], errors="coerce"
    ).fillna(0).astype(int)

    df["Capital cost storage (€/MWh)"] = pd.to_numeric(
        df["Capital cost storage (€/MWh)"], errors="coerce"
    ).fillna(0.1).astype(float)

    df["Capital cost inverter (€/MW)"] = pd.to_numeric(
        df["Capital cost inverter (€/MW)"], errors="coerce"
    ).fillna(0.1).astype(float)

    # -----------------------------
    # Crear componentes
    # -----------------------------
    for n in range(len(df)):
        location = df.loc[n, "LOCATION"]

        if pd.isna(location):
            continue

        p_nom_base = df.loc[n, "Rated active power (MW)"]
        max_hours = df.loc[n, "Max hours at rated active power (h)"]

        if pd.isna(p_nom_base) or pd.isna(max_hours):
            continue

        p_nom_base = float(p_nom_base)
        max_hours = float(max_hours)
        e_nom_base = p_nom_base * max_hours

        eta_store = float(df.loc[n, "efficiency_store (p.u)"])
        eta_dispatch = float(df.loc[n, "efficiency_dispatch (p.u)"])
        standing_loss = float(df.loc[n, "standing_loss (%/h)"])
        e_cyclic = bool(df.loc[n, "cyclic SOC (0/1)"])
        marginal_cost = float(df.loc[n, "marginal_cost (€/MWh)"])
        optimize_battery = int(df.loc[n, "Optimize battery (0/1)"]) == 1
        capex_storage = float(df.loc[n, "Capital cost storage (€/MWh)"])
        capex_inverter = float(df.loc[n, "Capital cost inverter (€/MW)"])

        initial_soc = float(df.loc[n, "initial SOC (%)"])
        if initial_soc > 1.0:
            initial_soc = initial_soc / 100.0

        e_initial = initial_soc * e_nom_base

        ac_bus = f"Bus_node_{location}"
        bat_bus = f"Bus_battery_{location}_{n}"

        store_name = f"BatteryStore_{location}_{n}"
        charge_link_name = f"BatteryCharge_{location}_{n}"
        discharge_link_name = f"BatteryDischarge_{location}_{n}"

        missing_capex = optimize_battery and (
            capex_storage == 0.0 or capex_inverter == 0.0
        )

        if missing_capex and not missing_capex_warning_shown:
            print("WARNING: Battery CAPEX missing. Using base values as maximum bounds.")
            missing_capex_warning_shown = True

        # Bus interno de batería
        if bat_bus not in grid.buses.index:
            grid.add(
                "Bus",
                bat_bus,
                carrier="AC",
            )

        # -----------------------------
        # Store
        # -----------------------------
        store_kwargs = dict(
            bus=bat_bus,
            e_initial=e_initial,
            e_cyclic=e_cyclic,
            standing_loss=standing_loss,
            marginal_cost=marginal_cost,
            capital_cost=capex_storage,
            carrier="AC",
        )

        if optimize_battery:
            store_kwargs.update(
                e_nom_extendable=True,
                e_nom_min=0.0,
            )

            if missing_capex:
                store_kwargs.update(
                    e_nom_max=e_nom_base
                )
        else:
            store_kwargs.update(
                e_nom=e_nom_base,
            )

        grid.add(
            "Store",
            store_name,
            **store_kwargs
        )

        # -----------------------------
        # Link de carga
        # -----------------------------
        charge_kwargs = dict(
            bus0=ac_bus,
            bus1=bat_bus,
            efficiency=eta_store,
            marginal_cost=0.0,
            capital_cost=capex_inverter,
            carrier="AC",
        )

        if optimize_battery:
            charge_kwargs.update(
                p_nom_extendable=True,
                p_nom_min=0.0,
            )

            if missing_capex:
                charge_kwargs.update(
                    p_nom_max=p_nom_base
                )
        else:
            charge_kwargs.update(
                p_nom=p_nom_base,
            )

        grid.add(
            "Link",
            charge_link_name,
            **charge_kwargs
        )

        # -----------------------------
        # Link de descarga
        # -----------------------------
        discharge_kwargs = dict(
            bus0=bat_bus,
            bus1=ac_bus,
            efficiency=eta_dispatch,
            marginal_cost=0.0,
            capital_cost=capex_inverter,
            carrier="AC",
        )

        if optimize_battery:
            discharge_kwargs.update(
                p_nom_extendable=True,
                p_nom_min=0.0,
            )

            if missing_capex:
                discharge_kwargs.update(
                    p_nom_max=p_nom_base
                )
        else:
            discharge_kwargs.update(
                p_nom=p_nom_base,
            )

        grid.add(
            "Link",
            discharge_link_name,
            **discharge_kwargs
        )

        # Guardamos info para restricciones
        battery_specs.append({
            "store_name": store_name,
            "charge_link_name": charge_link_name,
            "discharge_link_name": discharge_link_name,
            "max_hours": max_hours,
            "optimize_battery": optimize_battery,
        })

    return battery_specs

