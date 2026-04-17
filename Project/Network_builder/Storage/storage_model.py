import pypsa
import pandas as pd
import numpy as np



def add_storage_as_store_links(
    df_SYS_settings: pd.DataFrame,
    grid: pypsa.Network,
    df_StorageUnit: pd.DataFrame,
    CRF
) -> list[dict]:
    """
    Añade baterías como Store + Link a la red.

    Modos admitidos en 'Optimization mode':
        - Fixed
        - Optimize MWh
        - Optimize MW
        - Optimize both
    """
    simulation_duration = df_SYS_settings.loc[3, "SYSTEM PARAMETERS"]

    df = df_StorageUnit.copy()
    battery_specs = []
    missing_capex_warning_shown = False

    # -----------------------------
    # Conversión y valores por defecto
    # -----------------------------
    df["Efficiency store (p.u)"] = pd.to_numeric(
        df["Efficiency store (p.u)"], errors="coerce"
    ).fillna(0.95).astype(float)

    df["Efficiency dispatch (p.u)"] = pd.to_numeric(
        df["Efficiency dispatch (p.u)"], errors="coerce"
    ).fillna(0.95).astype(float)

    df["Standing loss (%/h)"] = pd.to_numeric(
        df["Standing loss (%/h)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["Cyclic SOC (0/1)"] = pd.to_numeric(
        df["Cyclic SOC (0/1)"], errors="coerce"
    ).fillna(1).astype(int)

    df["Initial SOC (%)"] = pd.to_numeric(
        df["Initial SOC (%)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["Marginal cost (€/MWh)"] = pd.to_numeric(
        df["Marginal cost (€/MWh)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["Rated active power (MW)"] = pd.to_numeric(
        df["Rated active power (MW)"], errors="coerce"
    )

    df["Capacity (MWh)"] = pd.to_numeric(
        df["Capacity (MWh)"], errors="coerce"
    )

    df["Investment cost storage (€/MWh)"] = pd.to_numeric(
        df["Investment cost storage (€/MWh)"], errors="coerce"
    ).fillna(0.0).astype(float)

    df["Investment cost inverter (€/MW)"] = pd.to_numeric(
        df["Investment cost inverter (€/MW)"], errors="coerce"
    ).fillna(0.0).astype(float)

    if "Optimization mode" not in df.columns:
        df["Optimization mode"] = "Fixed"

    df["Optimization mode"] = (
        df["Optimization mode"]
        .fillna("Fixed")
        .astype(str)
        .str.strip()
    )

    valid_modes = {"Fixed", "Optimize MWh", "Optimize MW", "Optimize both"}

    # -----------------------------
    # Crear componentes
    # -----------------------------
    for n in range(len(df)):
        location = df.loc[n, "LOCATION"]

        if pd.isna(location):
            continue

        optimization_mode = df.loc[n, "Optimization mode"]
        if optimization_mode not in valid_modes:
            raise ValueError(
                f"Invalid optimization mode in row {n}: {optimization_mode}"
            )

        optimize_e = optimization_mode in {"Optimize MWh", "Optimize both"}
        optimize_p = optimization_mode in {"Optimize MW", "Optimize both"}

        p_nom_base = df.loc[n, "Rated active power (MW)"]
        e_nom_base = df.loc[n, "Capacity (MWh)"]

        print(f"Row {n} - location={location}")
        print(f"p_nom_base={p_nom_base}, e_nom_base={e_nom_base}, mode={optimization_mode}")

        # -----------------------------
        # Validaciones según modo
        # -----------------------------
        if not optimize_p and pd.isna(p_nom_base):
            raise ValueError(
                f"Row {n}: 'Rated active power (MW)' is required for mode '{optimization_mode}'."
            )

        if not optimize_e and pd.isna(e_nom_base):
            raise ValueError(
                f"Row {n}: 'Capacity (MWh)' is required for mode '{optimization_mode}'."
            )

        # Convertir a float si existen
        p_nom_base = float(p_nom_base) if pd.notna(p_nom_base) else None
        e_nom_base = float(e_nom_base) if pd.notna(e_nom_base) else None

        eta_store = float(df.loc[n, "Efficiency store (p.u)"])
        eta_dispatch = float(df.loc[n, "Efficiency dispatch (p.u)"])
        standing_loss = float(df.loc[n, "Standing loss (%/h)"])
        e_cyclic = bool(df.loc[n, "Cyclic SOC (0/1)"])
        marginal_cost = float(df.loc[n, "Marginal cost (€/MWh)"])
        capex_storage = float(df.loc[n, "Investment cost storage (€/MWh)"])
        capex_inverter = float(df.loc[n, "Investment cost inverter (€/MW)"])

        initial_soc = float(df.loc[n, "Initial SOC (%)"])
        if initial_soc > 1.0:
            initial_soc = initial_soc / 100.0

        if not 0.0 <= initial_soc <= 1.0:
            raise ValueError(f"Row {n}: initial SOC out of range.")

        # e_initial solo tiene sentido si hay capacidad fija conocida
        if e_nom_base is not None:
            e_initial = initial_soc * e_nom_base
        else:
            e_initial = 0.0

        ac_bus = f"Bus_node_{location}"
        bat_bus = f"Bus_battery_{location}_{n}"

        store_name = f"BatteryStore_{location}_{n}"
        charge_link_name = f"BatteryCharge_{location}_{n}"
        discharge_link_name = f"BatteryDischarge_{location}_{n}"

        if bat_bus not in grid.buses.index:
            grid.add(
                "Bus",
                bat_bus,
                carrier="AC",
            )

        # Si falta CAPEX en una variable extendable, usar base como cota máxima
        missing_storage_capex = optimize_e and capex_storage == 0.0
        missing_inverter_capex = optimize_p and capex_inverter == 0.0

        if (missing_storage_capex or missing_inverter_capex) and not missing_capex_warning_shown:
            print(
                "WARNING: Battery CAPEX missing for an extendable component. "
                "Using base values as maximum bounds when available."
            )
            missing_capex_warning_shown = True

        # -----------------------------
        # Store
        # -----------------------------
        store_kwargs = dict(
            bus=bat_bus,
            e_initial=e_initial,
            e_cyclic=e_cyclic,
            standing_loss=standing_loss,
            marginal_cost=marginal_cost,
            capital_cost=capex_storage*CRF/365*simulation_duration,
            carrier="AC",
        )

        if optimize_e:
            store_kwargs.update(
                e_nom_extendable=True,
                e_nom_min=0.0,
            )
            if missing_storage_capex and e_nom_base is not None:
                store_kwargs.update(e_nom_max=e_nom_base)
        else:
            store_kwargs.update(
                e_nom=e_nom_base,
            )

        grid.add("Store", store_name, **store_kwargs)

        # -----------------------------
        # Link de carga
        # -----------------------------
        charge_kwargs = dict(
            bus0=ac_bus,
            bus1=bat_bus,
            efficiency=eta_store,
            marginal_cost=0.0,
            capital_cost=capex_inverter*CRF/365*simulation_duration,
            carrier="AC",
        )

        if optimize_p:
            charge_kwargs.update(
                p_nom_extendable=True,
                p_nom_min=0.0,
            )
            if missing_inverter_capex and p_nom_base is not None:
                charge_kwargs.update(p_nom_max=p_nom_base)
        else:
            charge_kwargs.update(
                p_nom=p_nom_base,
            )

        grid.add("Link", charge_link_name, **charge_kwargs)

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

        if optimize_p:
            discharge_kwargs.update(
                p_nom_extendable=True,
                p_nom_min=0.0,
            )
            if missing_inverter_capex and p_nom_base is not None:
                discharge_kwargs.update(p_nom_max=p_nom_base)
        else:
            discharge_kwargs.update(
                p_nom=p_nom_base,
            )

        grid.add("Link", discharge_link_name, **discharge_kwargs)

        battery_specs.append({
            "store_name": store_name,
            "charge_link_name": charge_link_name,
            "discharge_link_name": discharge_link_name,
            "optimize_p": optimize_p,
            "optimize_e": optimize_e,
            "optimization_mode": optimization_mode,
        })

    return battery_specs