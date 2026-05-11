import pypsa
import pandas as pd
import numpy as np
from Network_builder.Storage.runoff4hydro import get_hydro_inflow_node




def add_storage_as_store_links(
    df_SYS_settings: pd.DataFrame,
    grid: pypsa.Network,
    df_StorageUnit: pd.DataFrame,
    CRF,
    df_hydro_inflow_scaled: pd.DataFrame,
    initial_soc_fraction: float,
) -> list[dict]:
    
    """
    Añade baterías como Store + Link a la red.

    Modos admitidos en 'Optimization mode':
        - Fixed
        - Optimize MWh
        - Optimize MW
        - Optimize both
    """
    params = df_SYS_settings["SYSTEM PARAMETERS"]
    simulation_duration = params["Simulation duration (days)"]
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
        location = str(df.loc[n, "LOCATION"])

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
        carrier = str(df.loc[n, "Carrier"])
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

        ac_bus = f"Bus.{location}"
        bat_bus = f"Bus_battery_{location}_{n}"
        
        store_name = f"{carrier}_{location}_{n}"

        if carrier!="PHS" and carrier!="hydro":
            charge_link_name = f"BatteryCharge_{location}_{n}"
            discharge_link_name = f"BatteryDischarge_{location}_{n}"
        else:
            charge_link_name = f"{carrier}_Charge_{location}_{n}"
            discharge_link_name = f"{carrier}_Discharge_{location}_{n}"

        if bat_bus not in grid.buses.index:
            grid.add(
                "Bus",
                bat_bus,
                carrier=carrier,
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
        if carrier == "hydro":
            initial_soc = initial_soc_fraction
            e_initial = initial_soc * e_nom_base

        store_kwargs = dict(
            bus=bat_bus,
            e_initial=e_initial,
            e_cyclic=e_cyclic,
            standing_loss=standing_loss,
            marginal_cost=marginal_cost,
            capital_cost=capex_storage*CRF/365*simulation_duration,
            carrier=carrier,
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
        

        if carrier == "hydro":
            gen_name = f"HydroInflow_{location}_{n}"

            inflow = get_hydro_inflow_node(df_hydro_inflow_scaled, location)
            inflow = pd.to_numeric(inflow, errors="coerce")
            inflow.index = pd.to_datetime(inflow.index)
            inflow = inflow.reindex(grid.snapshots).fillna(0.0)

            p_nom_inflow = inflow.max()

            if p_nom_inflow > 0:
                profile = inflow / p_nom_inflow

                grid.add(
                    "Generator",
                    gen_name,
                    bus=bat_bus,
                    carrier="hydro inflow",
                    p_nom=p_nom_inflow,
                    marginal_cost=0.0,
                    p_min_pu=0.0,
                    p_max_pu=1.0,
                )

                grid.generators_t.p_max_pu.loc[:, gen_name] = profile
            else:
                print(f"Aviso: inflow máximo cero para {location}. No se añade generador de inflow.")



        annualized_inverter_cost = capex_inverter * CRF / 365 * simulation_duration
        link_capital_cost = 0.5 * annualized_inverter_cost
        # -----------------------------
        # Link de carga
        # -----------------------------
        charge_kwargs = dict(
            bus0=ac_bus,
            bus1=bat_bus,
            efficiency=eta_store,
            marginal_cost=0.0,
            capital_cost=link_capital_cost,
            carrier=carrier,
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
            capital_cost=link_capital_cost,
            carrier=carrier,
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