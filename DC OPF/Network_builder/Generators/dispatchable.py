import math
import pypsa
import pandas as pd


def add_dispatchable_generators(
    grid: pypsa.Network,
    df_Gen_Dispatchable: pd.DataFrame
) -> None:

    df = df_Gen_Dispatchable.copy()

    numeric_cols = [
        "Rated active power (MW)",
        "Pmin (MW)",
        "Ramp limit up (p.u)",
        "Ramp limit down (p.u)",
        "Ramp limit start up (p.u)",
        "Ramp limit shut down (p.u)",
        "Start up cost (€)",
        "Shut down cost (€)",
        "Stand by cost (€/h)",
        "Min up time (h)",
        "Min down time (h)",
        "Up time before (h)",
        "Down time before (h)",
        "Initial power (MW)",
        "€/MW²h",
        "€/MWh",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Defaults razonables
    df["Pmin (MW)"] = df["Pmin (MW)"].fillna(0.0)
    df["Ramp limit start up (p.u)"] = df["Ramp limit start up (p.u)"].fillna(1.0)
    df["Ramp limit shut down (p.u)"] = df["Ramp limit shut down (p.u)"].fillna(1.0)
    df["Start up cost (€)"] = df["Start up cost (€)"].fillna(0.0)
    df["Shut down cost (€)"] = df["Shut down cost (€)"].fillna(0.0)
    df["Stand by cost (€/h)"] = df["Stand by cost (€/h)"].fillna(0.0)
    df["Min up time (h)"] = df["Min up time (h)"].fillna(0)
    df["Min down time (h)"] = df["Min down time (h)"].fillna(0)
    df["€/MWh"] = df["€/MWh"].fillna(0.0)

    for n in range(len(df)):
        Pmax = df.loc[n, "Rated active power (MW)"]
        location = df.loc[n, "GENERATOR LOCATION"]

        if pd.isna(Pmax) or pd.isna(location) or Pmax <= 0:
            continue

        Pmax = float(Pmax)
        location = int(location)

        Pmin = float(df.loc[n, "Pmin (MW)"])
        marginal_cost = float(df.loc[n, "€/MWh"])

        ramp_limit_up = df.loc[n, "Ramp limit up (p.u)"]
        ramp_limit_down = df.loc[n, "Ramp limit down (p.u)"]
        ramp_limit_start_up = float(df.loc[n, "Ramp limit start up (p.u)"])
        ramp_limit_shut_down = float(df.loc[n, "Ramp limit shut down (p.u)"])

        min_up_time = int(df.loc[n, "Min up time (h)"])
        min_down_time = int(df.loc[n, "Min down time (h)"])

        up_time_before_raw = df.loc[n, "Up time before (h)"]
        down_time_before_raw = df.loc[n, "Down time before (h)"]

        start_up_cost = float(df.loc[n, "Start up cost (€)"])
        shut_down_cost = float(df.loc[n, "Shut down cost (€)"])
        stand_by_cost = float(df.loc[n, "Stand by cost (€/h)"])

        p_init_raw = df.loc[n, "Initial power (MW)"]

        # Validaciones simples
        if Pmin < 0:
            Pmin = 0.0
        if Pmin > Pmax:
            Pmin = Pmax

        p_min_pu = Pmin / Pmax if Pmax > 0 else 0.0

        # Lógica por defecto para historia previa:
        # - si el usuario no pone nada, asumimos unidad encendida y libre
        #   respecto a min_up_time al inicio del horizonte
        if pd.isna(up_time_before_raw) and pd.isna(down_time_before_raw):
            up_time_before = min_up_time
            down_time_before = 0
        else:
            up_time_before = 0 if pd.isna(up_time_before_raw) else int(up_time_before_raw)
            down_time_before = 0 if pd.isna(down_time_before_raw) else int(down_time_before_raw)

        # Evitar inconsistencias obvias
        if up_time_before > 0 and down_time_before > 0:
            # Priorizamos "encendido antes del horizonte"
            down_time_before = 0

        # -----------------------------
        # Validación robusta de p_init
        # -----------------------------
        add_kwargs = {}

        if not pd.isna(p_init_raw):
            p_init = float(p_init_raw)

            # 1) Acotar entre 0 y Pmax
            p_init = max(0.0, min(p_init, Pmax))

            # 2) Si el historial indica unidad apagada antes del horizonte,
            #    la potencia inicial debe ser 0
            if down_time_before > 0:
                p_init = 0.0

            # 3) Si el historial indica unidad encendida,
            #    una potencia positiva por debajo de Pmin no es consistente
            elif up_time_before > 0:
                if 0.0 < p_init < Pmin:
                    p_init = Pmin

            # 4) Caso ambiguo: no hay señal clara de ON/OFF previa
            #    Si hay potencia positiva pero menor que Pmin, la corregimos a Pmin
            else:
                if 0.0 < p_init < Pmin:
                    p_init = Pmin

            add_kwargs["p_init"] = p_init

        # Solo añadimos rampas si el usuario las definió
        if not pd.isna(ramp_limit_up):
            add_kwargs["ramp_limit_up"] = float(ramp_limit_up)

        if not pd.isna(ramp_limit_down):
            add_kwargs["ramp_limit_down"] = float(ramp_limit_down)

        add_kwargs["ramp_limit_start_up"] = ramp_limit_start_up
        add_kwargs["ramp_limit_shut_down"] = ramp_limit_shut_down

        grid.add(
            "Generator",
            f"DispatchGen{location}_{n}",
            bus=f"Bus_node_{location}",
            p_nom=Pmax,
            p_min_pu=p_min_pu,
            p_max_pu=1.0,
            marginal_cost=marginal_cost,
            start_up_cost=start_up_cost,
            shut_down_cost=shut_down_cost,
            stand_by_cost=stand_by_cost,
            min_up_time=min_up_time,
            min_down_time=min_down_time,
            up_time_before=up_time_before,
            down_time_before=down_time_before,
            committable=True,
            carrier="AC",
            **add_kwargs,
        )