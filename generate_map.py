import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import contextily as ctx
from pyproj import Transformer

df = pd.read_csv("surveyed_locations_electric_wheelchair.csv")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
x, y = transformer.transform(df["longitude"].values, df["latitude"].values)

fig, ax = plt.subplots(figsize=(10, 10))

ax.scatter(x, y, color="#008080", s=60, zorder=5, linewidths=0.5, edgecolors="white")

pad_x = (max(x) - min(x)) * 0.18
pad_y = (max(y) - min(y)) * 0.18
ax.set_xlim(min(x) - pad_x, max(x) + pad_x)
ax.set_ylim(min(y) - pad_y, max(y) + pad_y)

ctx.add_basemap(ax, crs="EPSG:3857", source=ctx.providers.CartoDB.Positron, zoom=13)

legend_handle = mpatches.Patch(color="#008080", label="Surveyed Location")
ax.legend(
    handles=[legend_handle],
    loc="lower right",
    fontsize=11,
    framealpha=0.9,
    edgecolor="#cccccc",
    fancybox=False,
)

ax.set_axis_off()
fig.patch.set_facecolor("white")

plt.tight_layout(pad=0)
plt.savefig("docs/images/surveyed_locations_map.png", dpi=180, bbox_inches="tight", facecolor="white")
print("Saved to docs/images/surveyed_locations_map.png")
