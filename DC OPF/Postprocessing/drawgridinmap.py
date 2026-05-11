
import pypsa
import matplotlib.pyplot as plt
import cartopy.crs as ccrs
import cartopy.feature as cfeature
import pandas as pd


def create_buses_with_drawing_names(grid: pypsa.Network,
                                    df_Net_Buses: pd.DataFrame) -> pd.DataFrame:
    
    buses_new = grid.buses.copy()
    new_index = []
    c=0
    for i in buses_new.index:
        drawing_name = df_Net_Buses.loc[c, "Bus name"]

        c=c+1
        
        if pd.notna(drawing_name):
            new_index.append(drawing_name)

        else:
            new_index.append(i)   # ← mantiene el nombre original del bus

    buses_new.index = new_index
    
    return buses_new



def drawrealgrid(grid: pypsa.Network, df_Net_Buses: pd.DataFrame, filename):

    fig, ax = plt.subplots(
        figsize=(8, 6),
        subplot_kw={"projection": ccrs.PlateCarree()}
    )

    lon_pad = 0.3
    lat_pad = 0.3
    lon_min = grid.buses.x.min()-lon_pad
    lon_max = grid.buses.x.max()+lon_pad

    lat_min = grid.buses.y.min()-lat_pad
    lat_max = grid.buses.y.max()+lat_pad

    ax.set_extent([lon_min, lon_max, lat_min, lat_max], crs=ccrs.PlateCarree())

    ax.add_feature(cfeature.LAND, facecolor="lightgray")
    ax.add_feature(cfeature.OCEAN, facecolor="aliceblue")
    ax.add_feature(cfeature.COASTLINE)
    ax.add_feature(cfeature.BORDERS, linestyle=":")
    
    # ----------------------
    # 1. Dibujar conexiones entre buses
    # ----------------------
    # Dibujar líneas AC
    for _, line in grid.lines.iterrows():
        b0 = grid.buses.loc[line.bus0]
        b1 = grid.buses.loc[line.bus1]

        ax.plot(
            [b0.x, b1.x],
            [b0.y, b1.y],
            color="black",
            linewidth=2,
            transform=ccrs.PlateCarree(),
            zorder=2
        )


    # Dibujar links, por ejemplo interconexiones/PCCs
    for _, link in grid.links.iterrows():
        b0 = grid.buses.loc[link.bus0]
        b1 = grid.buses.loc[link.bus1]

        ax.plot(
            [b0.x, b1.x],
            [b0.y, b1.y],
            color="red",
            linewidth=2,
            linestyle="--",
            transform=ccrs.PlateCarree(),
            zorder=2
        )

    # ----------------------
    # 2. Dibujar buses
    # ----------------------
    ax.scatter(
        grid.buses.x,
        grid.buses.y,
        s=150,
        color="red",
        edgecolors="black",
        transform=ccrs.PlateCarree(),
        zorder=3
    )

    
    

    # Etiquetas
    """
    buses_new = create_buses_with_drawing_names(grid, df_Net_Buses)
    for name, bus in buses_new.iterrows():
        ax.text(
            bus.x + 0.1,
            bus.y + 0.1,
            name,
            transform=ccrs.PlateCarree(),
            fontsize=10,
            zorder=4
        )
    """
    
    plt.title("Electric grid (PyPSA + Cartopy)")
    plt.savefig(filename, dpi=300, bbox_inches="tight")
    plt.close(fig)

