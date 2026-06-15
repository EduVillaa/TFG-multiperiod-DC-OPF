import re
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import cartopy.crs as ccrs
import cartopy.feature as cfeature
import matplotlib.tri as mtri


def build_node_curtailment_timeseries(
    curtailment_df: pd.DataFrame,
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
) -> pd.DataFrame | None:
    """
    Construye un DataFrame temporal de curtailment agregado por nodo
    a partir de columnas tipo:
        PV_ES0 0_curtailment
        Wind_ES0 0_curtailment
        ...

    Devuelve un DataFrame con:
        index   -> snapshots
        columns -> nodos
        values  -> curtailment agregado PV + Wind por nodo
    """

    df = curtailment_df.copy()

    if snapshot_col in df.columns:
        df[snapshot_col] = pd.to_datetime(df[snapshot_col])
        df = df.set_index(snapshot_col)
    else:
        df.index = pd.to_datetime(df.index)

    tech_pattern = "|".join(technologies)
    pattern = re.compile(rf"^({tech_pattern})_(.+)_curtailment$")

    node_data = {}

    for col in df.columns:
        match = pattern.match(col)
        if match is None:
            continue

        node = match.group(2)

        if node not in node_data:
            node_data[node] = df[col].astype(float)
        else:
            node_data[node] = node_data[node].add(df[col].astype(float), fill_value=0)

    if not node_data:
        print("En build_node_curtailment_timeseries: not node_data")
        return None

    node_curt = pd.DataFrame(node_data, index=df.index)
    node_curt = node_curt.replace([np.inf, -np.inf], np.nan).fillna(0)
    node_curt = node_curt.loc[:, node_curt.sum(axis=0) > 0]

    if node_curt.empty:
        print("En build_node_curtailment_timeseries: node_curt.empty:")
        return None

    return node_curt


def build_node_curtailment_map_df(
    grid,
    curtailment_df: pd.DataFrame,
    agg: str = "sum",
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
    top_n_nodes: int | None = None,
) -> pd.DataFrame | None:
    """
    Devuelve un DataFrame con:
        node, lon, lat, value

    donde 'value' es el curtailment agregado por nodo.
    """

    node_ts = build_node_curtailment_timeseries(
        curtailment_df=curtailment_df,
        technologies=technologies,
        snapshot_col=snapshot_col,
    )

    if node_ts is None or node_ts.empty:
        print("En build_node_curtailment_map_df: node_ts is None or node_ts.empty")
        return None

    if agg == "sum":
        values = node_ts.sum(axis=0)
    elif agg == "mean":
        values = node_ts.mean(axis=0)
    elif agg == "max":
        values = node_ts.max(axis=0)
    else:
        raise ValueError("agg debe ser 'sum', 'mean' o 'max'")

    values = values[values > 0]

    if values.empty:
        print("En build_node_curtailment_map_df: values.empty")
        return None

    # Coordenadas desde grid.buses
    if grid.buses.empty or "x" not in grid.buses.columns or "y" not in grid.buses.columns:
        print("En build_node_curtailment_map_df: grid.buses.empty or 'x' not in grid.buses.columns or 'y' not in grid.buses.columns")
        return None

    coords = grid.buses[["x", "y"]].rename(columns={"x": "lon", "y": "lat"}).copy()

    # Adaptar nombres: Bus.ES0 10 -> ES0 10
    coords.index = (
    coords.index
    .astype(str)
    .str.replace("Bus.", "", regex=False)
    .str.strip())


    map_df = pd.DataFrame({
        "node": values.index,
        "value": values.values,
    }).set_index("node")



    map_df = map_df.join(coords, how="inner")
    map_df = map_df.dropna(subset=["lon", "lat"])

    if map_df.empty:

        return None

    map_df = map_df.sort_values("value", ascending=False)

    if top_n_nodes is not None:
        map_df = map_df.head(top_n_nodes)

    map_df.index.name = "node"

    return map_df.reset_index()





def plot_curtailment_geo_heatmap_nodes(
    grid,
    curtailment_df: pd.DataFrame,
    agg: str = "sum",
    top_n_nodes: int | None = 15,
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
    annotate: bool = True,
):
    """
    Mapa geográfico del curtailment por nodo.
    El color y el tamaño del nodo representan el curtailment.
    """

    map_df = build_node_curtailment_map_df(
        grid=grid,
        curtailment_df=curtailment_df,
        agg=agg,
        technologies=technologies,
        snapshot_col=snapshot_col,
        top_n_nodes=top_n_nodes,
    )

    if map_df is None or map_df.empty:
        return None

    lons = map_df["lon"].values
    lats = map_df["lat"].values
    vals = map_df["value"].values
    labels = map_df["node"].values

    # Escalado del tamaño de puntos
    vmin = vals.min()
    vmax = vals.max()

    if np.isclose(vmin, vmax):
        sizes = np.full_like(vals, 180, dtype=float)
    else:
        sizes = 80 + 320 * (vals - vmin) / (vmax - vmin)

    fig = plt.figure(figsize=(12, 9))
    ax = plt.axes(projection=ccrs.PlateCarree())

    # Extensión aproximada península ibérica
    ax.set_extent([-10.5, 4.5, 35.0, 44.8], crs=ccrs.PlateCarree())

    # Fondo del mapa
    ax.add_feature(cfeature.LAND, facecolor="#f2f2f2")
    ax.add_feature(cfeature.OCEAN, facecolor="#dbeaf7")
    ax.add_feature(cfeature.COASTLINE, linewidth=1.0)
    ax.add_feature(cfeature.BORDERS, linestyle=":", linewidth=1.0)
    #ax.add_feature(cfeature.LAKES, alpha=0.4)
    #ax.add_feature(cfeature.RIVERS, alpha=0.3)

    # Scatter tipo heatmap
    sc = ax.scatter(
        lons,
        lats,
        c=vals,
        s=sizes,
        cmap="YlOrRd",
        edgecolors="black",
        linewidths=0.7,
        alpha=0.95,
        transform=ccrs.PlateCarree(),
        zorder=5,
    )

    if annotate:
        for lon, lat, label in zip(lons, lats, labels):
            ax.text(
                lon + 0.08,
                lat + 0.05,
                label,
                fontsize=8,
                transform=ccrs.PlateCarree(),
                zorder=6,
            )

    cbar = fig.colorbar(sc, ax=ax, shrink=0.8, pad=0.03)
    if agg == "sum":
        cbar.set_label("Curtailment total [MWh]")
        title_agg = "total"
    elif agg == "mean":
        cbar.set_label("Curtailment medio [MW]")
        title_agg = "mean"
    else:
        cbar.set_label("Curtailment máximo [MW]")
        title_agg = "maximum"

    ax.set_title(
        f"PV + Wind curtailment map ({title_agg}) - Top {len(map_df)} nodes",
        fontsize=13
    )

    fig.tight_layout()
    return fig





def plot_curtailment_geo_heatmap_interpolated(
    grid,
    curtailment_df: pd.DataFrame,
    agg: str = "sum",
    top_n_nodes: int | None = None,
    technologies: tuple[str, ...] = ("PV", "Wind"),
    snapshot_col: str = "snapshot",
    show_nodes: bool = True,
):
    """
    Mapa geográfico con heatmap interpolado a partir de nodos.
    OJO: visualmente atractivo, pero puede sugerir valores entre nodos.
    """

    map_df = build_node_curtailment_map_df(
        grid=grid,
        curtailment_df=curtailment_df,
        agg=agg,
        technologies=technologies,
        snapshot_col=snapshot_col,
        top_n_nodes=top_n_nodes,
    )

    if map_df is None or len(map_df) < 3:
        print("Ejecución de plot_curtailment_geo_heatmap_interpolated")
        print("map_df is None or len(map_df) < 3")
        return None
    

    lons = map_df["lon"].values
    lats = map_df["lat"].values
    vals = map_df["value"].values

    fig = plt.figure(figsize=(12, 9))
    ax = plt.axes(projection=ccrs.PlateCarree())
    ax.set_extent([-10.5, 4.5, 35.0, 44.8], crs=ccrs.PlateCarree())

    ax.add_feature(cfeature.LAND, facecolor="#f2f2f2")
    ax.add_feature(cfeature.OCEAN, facecolor="#dbeaf7")
    ax.add_feature(cfeature.COASTLINE, linewidth=1.0)
    ax.add_feature(cfeature.BORDERS, linestyle=":", linewidth=1.0)

    triang = mtri.Triangulation(lons, lats)

    contour = ax.tricontourf(
        triang,
        vals,
        levels=15,
        cmap="YlOrRd",
        alpha=0.85,
        transform=ccrs.PlateCarree(),
        zorder=3,
    )

    if show_nodes:
        ax.scatter(
            lons,
            lats,
            c=vals,
            cmap="YlOrRd",
            edgecolors="black",
            s=50,
            transform=ccrs.PlateCarree(),
            zorder=5,
        )

    cbar = fig.colorbar(contour, ax=ax, shrink=0.8, pad=0.03)
    if agg == "sum":
        cbar.set_label("Curtailment total [MWh]")
        title_agg = "total"
    elif agg == "mean":
        cbar.set_label("Curtailment medio [MW]")
        title_agg = "mean"
    else:
        cbar.set_label("Curtailment máximo [MW]")
        title_agg = "maximum"

    ax.set_title(f"Interpolated PV + Wind curtailment heatmap ({title_agg})")
    fig.tight_layout()

    return fig