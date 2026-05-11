import pypsa
import pandas as pd

# Cargar red
n = pypsa.Network("base_s_50_elec.nc")

# ============================================================
# 1. BUSES
# ============================================================

buses = n.buses.copy()

# Si quieres solo buses AC:
buses_ac = n.buses[n.buses.carrier == "AC"].copy()

# ============================================================
# 2. LÍNEAS
# ============================================================

lines = n.lines.copy()

# ============================================================
# 3. GENERADORES CONVENCIONALES Y RENOVABLES
# ============================================================

generators = n.generators.copy()

# Series temporales de generadores
generators_p_max_pu = n.generators_t.p_max_pu.copy()
generators_p_min_pu = n.generators_t.p_min_pu.copy()
generators_p_set = n.generators_t.p_set.copy()
generators_marginal_cost = n.generators_t.marginal_cost.copy()

# ============================================================
# 4. DEMANDA
# ============================================================

loads = n.loads.copy()
loads_timeseries = n.loads_t.p_set.copy()

# ============================================================
# 5. STORAGE UNITS
# Aquí suele estar la hidráulica de embalse / PHS
# ============================================================

storage_units = n.storage_units.copy()

# Series temporales de storage units
storage_units_inflow = n.storage_units_t.inflow.copy()
storage_units_p_set = n.storage_units_t.p_set.copy()
storage_units_p_max_pu = n.storage_units_t.p_max_pu.copy()
storage_units_p_min_pu = n.storage_units_t.p_min_pu.copy()

# Hidráulica de embalse si está como storage_unit
hydro_storage_units = storage_units[
    storage_units.carrier.astype(str).str.contains("hydro|reservoir|PHS", case=False, na=False)
].copy()

# ============================================================
# 6. STORES
# Baterías, depósitos de energía, embalses si están modelados como stores
# ============================================================

stores = n.stores.copy()

battery_energy = stores[
    stores.carrier.astype(str).str.contains("battery", case=False, na=False)
].copy()

hydro_stores = stores[
    stores.carrier.astype(str).str.contains("hydro|reservoir|water|PHS", case=False, na=False)
].copy()

# Series temporales de stores
stores_e_set = n.stores_t.e_set.copy()
stores_e_min_pu = n.stores_t.e_min_pu.copy()
stores_e_max_pu = n.stores_t.e_max_pu.copy()

# ============================================================
# 7. LINKS
# Carga/descarga de baterías, bombeo, turbinas hidráulicas, etc.
# ============================================================

links = n.links.copy()

battery_links = links[
    links.carrier.astype(str).str.contains("battery", case=False, na=False)
].copy()

battery_charge = battery_links[
    battery_links.index.astype(str).str.contains("charger|charge", case=False, na=False)
].copy()

battery_discharge = battery_links[
    battery_links.index.astype(str).str.contains("discharger|discharge", case=False, na=False)
].copy()

hydro_links = links[
    links.carrier.astype(str).str.contains("hydro|PHS|pump|reservoir|water", case=False, na=False)
].copy()

# Series temporales de links
links_p_set = n.links_t.p_set.copy()
links_p_max_pu = n.links_t.p_max_pu.copy()
links_p_min_pu = n.links_t.p_min_pu.copy()
links_marginal_cost = n.links_t.marginal_cost.copy()

# ============================================================
# 8. DIAGNÓSTICO POR PANTALLA
# ============================================================

print("=== Carriers en generators ===")
print(generators.carrier.value_counts())

print("\n=== Carriers en storage_units ===")
if len(storage_units) > 0:
    print(storage_units.carrier.value_counts())
else:
    print("No hay storage_units")

print("\n=== Carriers en stores ===")
if len(stores) > 0:
    print(stores.carrier.value_counts())
else:
    print("No hay stores")

print("\n=== Carriers en links ===")
if len(links) > 0:
    print(links.carrier.value_counts())
else:
    print("No hay links")

# ============================================================
# 9. EXPORTAR A EXCEL
# ============================================================

with pd.ExcelWriter("network_clean.xlsx") as writer:

    # Red
    buses.to_excel(writer, sheet_name="buses_all")
    buses_ac.to_excel(writer, sheet_name="buses_AC")
    lines.to_excel(writer, sheet_name="lines")

    # Generadores
    generators.to_excel(writer, sheet_name="generators")
    generators_p_max_pu.to_excel(writer, sheet_name="gen_p_max_pu")
    generators_p_min_pu.to_excel(writer, sheet_name="gen_p_min_pu")
    generators_p_set.to_excel(writer, sheet_name="gen_p_set")
    generators_marginal_cost.to_excel(writer, sheet_name="gen_marginal_cost")

    # Demanda
    loads.to_excel(writer, sheet_name="loads")
    loads_timeseries.to_excel(writer, sheet_name="loads_timeseries")

    # Storage units
    storage_units.to_excel(writer, sheet_name="storage_units")
    hydro_storage_units.to_excel(writer, sheet_name="hydro_storage_units")
    storage_units_inflow.to_excel(writer, sheet_name="storage_inflow")
    storage_units_p_set.to_excel(writer, sheet_name="storage_p_set")
    storage_units_p_max_pu.to_excel(writer, sheet_name="storage_p_max_pu")
    storage_units_p_min_pu.to_excel(writer, sheet_name="storage_p_min_pu")

    # Stores
    stores.to_excel(writer, sheet_name="stores")
    battery_energy.to_excel(writer, sheet_name="battery_energy")
    hydro_stores.to_excel(writer, sheet_name="hydro_stores")
    stores_e_set.to_excel(writer, sheet_name="stores_e_set")
    stores_e_min_pu.to_excel(writer, sheet_name="stores_e_min_pu")
    stores_e_max_pu.to_excel(writer, sheet_name="stores_e_max_pu")

    # Links
    links.to_excel(writer, sheet_name="links")
    battery_links.to_excel(writer, sheet_name="battery_links")
    battery_charge.to_excel(writer, sheet_name="battery_charge")
    battery_discharge.to_excel(writer, sheet_name="battery_discharge")
    hydro_links.to_excel(writer, sheet_name="hydro_links")
    links_p_set.to_excel(writer, sheet_name="links_p_set")
    links_p_max_pu.to_excel(writer, sheet_name="links_p_max_pu")
    links_p_min_pu.to_excel(writer, sheet_name="links_p_min_pu")
    links_marginal_cost.to_excel(writer, sheet_name="links_marginal_cost")

print("✅ Excel exportado como network_clean.xlsx")