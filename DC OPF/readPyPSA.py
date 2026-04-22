import pypsa
import pandas as pd

# Cargar red
n = pypsa.Network("base_s_50_elec.nc")

# =========================
# 1. BUSES (solo AC)
# =========================
buses = n.buses[n.buses.carrier == "AC"].copy()

# =========================
# 2. LÍNEAS
# =========================
lines = n.lines.copy()

# =========================
# 3. GENERADORES
# =========================
generators = n.generators.copy()

# =========================
# 4. DEMANDA (LOADS)
# =========================
loads = n.loads_t.p_set.copy()

# =========================
# 5. BATERÍAS
# =========================

# --- Energía (MWh) ---
stores = n.stores[n.stores.carrier == "battery"].copy()

# --- Potencia (MW) ---
links = n.links[n.links.carrier == "battery"].copy()

# Separar carga y descarga (opcional pero recomendable)
battery_charge = links[links.index.str.contains("charger", case=False)].copy()
battery_discharge = links[links.index.str.contains("discharger", case=False)].copy()

# =========================
# 6. EXPORTAR A EXCEL
# =========================

with pd.ExcelWriter("network_clean.xlsx") as writer:
    
    buses.to_excel(writer, sheet_name="buses")
    lines.to_excel(writer, sheet_name="lines")
    generators.to_excel(writer, sheet_name="generators")
    loads.to_excel(writer, sheet_name="loads_timeseries")
    stores.to_excel(writer, sheet_name="battery_energy")
    battery_charge.to_excel(writer, sheet_name="battery_charge")
    battery_discharge.to_excel(writer, sheet_name="battery_discharge")

print("✅ Excel exportado como network_clean.xlsx")