
import pandas as pd
from pathlib import Path

ruta = Path(r"C:\Users\eduavill\Desktop\Load profiles")
año = 2015
dfs = []  # ← aquí acumulamos los DataFrames

for i in range(1, 13):
    archivo = ruta / f"PERFF_{año}{i:02d}.txt"

    if i == 1:
        df = pd.read_csv(
            archivo,
            encoding="latin-1",
            sep=";",
            usecols=["AÑO", "MES", "DIA", "HORA",
                     "COEF. PERFIL A", "COEF. PERFIL B",
                     "COEF. PERFIL C", "COEF. PERFIL D"]
        )
    else:
        df = pd.read_csv(
            archivo,
            encoding="latin-1",
            sep=";",
            skiprows=1,
            usecols=[0, 1, 2, 3, 5, 6, 7, 8]
        )
        # opcional: forzar los mismos nombres de columnas
        df.columns = dfs[0].columns

    dfs.append(df)  # ← aquí se va acumulando

# concatenación final
df_final = pd.concat(dfs, ignore_index=True)


divisores = {
    "COEF. PERFIL A": df_final.loc[:, "COEF. PERFIL A"].sum(),
    "COEF. PERFIL B": df_final.loc[:, "COEF. PERFIL B"].sum(),
    "COEF. PERFIL C": df_final.loc[:, "COEF. PERFIL C"].sum(),
    "COEF. PERFIL D": df_final.loc[:, "COEF. PERFIL D"].sum(),
}

df_final = df_final.drop(df_final.columns[[0, 1, 2, 3]], axis=1)
df_final = df_final.div(divisores)

print(df_final.loc[:, "COEF. PERFIL A"].sum())
print(df_final.loc[:, "COEF. PERFIL B"].sum())
print(df_final.loc[:, "COEF. PERFIL C"].sum())
print(df_final.loc[:, "COEF. PERFIL D"].sum())
print(df_final)

#df_final.to_excel("load_profiles_2015.xlsx", index=False)


df = df.drop('columna', axis=1)
df = df.drop(['col1', 'col2', 'col3'], axis=1)
df = df.loc[:, ['col1', 'col2', 'col3']]
df = df.dropna(axis=1)