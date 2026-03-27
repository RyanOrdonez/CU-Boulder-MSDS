"""
Generate static map images for the Folium slides using matplotlib.
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np
import os

IMG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "images")
os.makedirs(IMG, exist_ok=True)

# Launch site coordinates
sites = {
    'CCAFS LC-40':  (28.5618, -80.5774),
    'CCAFS SLC-40': (28.5618, -80.5774),
    'KSC LC-39A':   (28.6080, -80.6043),
    'VAFB SLC-4E':  (34.6321, -120.6108),
}

# Load data for success/failure
df = pd.read_csv(
    "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/"
    "IBM-DS0321EN-SkillsNetwork/datasets/spacex_launch_dash.csv"
)

# ── Map 1: Launch sites overview with success/failure ─────────────────
fig, ax = plt.subplots(figsize=(14, 8))
ax.set_facecolor('#e6f0ff')

for site_name, (lat, lon) in sites.items():
    site_df = df[df['Launch Site'] == site_name] if site_name in df['Launch Site'].values else pd.DataFrame()
    successes = site_df[site_df['class'] == 1].shape[0] if len(site_df) > 0 else 0
    failures  = site_df[site_df['class'] == 0].shape[0] if len(site_df) > 0 else 0
    total = successes + failures
    rate = successes / total * 100 if total > 0 else 0

    ax.scatter(lon, lat, s=300, c='steelblue', edgecolors='navy', linewidth=2, zorder=5)
    ax.annotate(f'{site_name}\n{successes}S / {failures}F ({rate:.0f}%)',
                (lon, lat), textcoords="offset points", xytext=(15, 10),
                fontsize=10, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.9))

ax.set_xlabel('Longitude', fontsize=13)
ax.set_ylabel('Latitude', fontsize=13)
ax.set_title('SpaceX Launch Sites — Success vs Failure Overview', fontsize=16, fontweight='bold')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(IMG, "folium_launch_sites_overview.png"), dpi=150, bbox_inches='tight')
plt.close()
print("Saved folium_launch_sites_overview.png")

# ── Map 2: KSC LC-39A Proximity Analysis ─────────────────────────────
fig, ax = plt.subplots(figsize=(12, 8))
ax.set_facecolor('#e6f0ff')

ksc_lat, ksc_lon = 28.6080, -80.6043
ax.scatter(ksc_lon, ksc_lat, s=400, c='red', edgecolors='darkred', linewidth=2, zorder=5, marker='*')
ax.annotate('KSC LC-39A', (ksc_lon, ksc_lat), textcoords="offset points", xytext=(15, 15),
            fontsize=14, fontweight='bold', color='darkred')

# Proximity points
proximities = {
    'Coastline (~1.2 km)': (28.6168, -80.5928),
    'Highway US-1 (~7.1 km)': (28.5733, -80.5400),
    'Railway (~5.7 km)': (28.5733, -80.5467),
    'City: Melbourne, FL (~52 km)': (28.0836, -80.6081),
}

colors = ['blue', 'orange', 'green', 'purple']
for (label, (plat, plon)), c in zip(proximities.items(), colors):
    ax.scatter(plon, plat, s=200, c=c, edgecolors='black', linewidth=1.5, zorder=4)
    ax.plot([ksc_lon, plon], [ksc_lat, plat], '--', color=c, alpha=0.6, linewidth=2)
    ax.annotate(label, (plon, plat), textcoords="offset points", xytext=(10, -15),
                fontsize=10, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.9))

ax.set_xlabel('Longitude', fontsize=13)
ax.set_ylabel('Latitude', fontsize=13)
ax.set_title('KSC LC-39A — Proximity Analysis', fontsize=16, fontweight='bold')
ax.grid(True, alpha=0.3)
plt.tight_layout()
plt.savefig(os.path.join(IMG, "folium_ksc_proximity.png"), dpi=150, bbox_inches='tight')
plt.close()
print("Saved folium_ksc_proximity.png")

# ── Map 3: Success/Failure markers at each site ──────────────────────
fig, axes = plt.subplots(1, 3, figsize=(18, 6))
site_list = ['CCAFS SLC-40', 'KSC LC-39A', 'VAFB SLC-4E']

for ax, site in zip(axes, site_list):
    site_df = df[df['Launch Site'] == site]
    succ = site_df[site_df['class'] == 1]
    fail = site_df[site_df['class'] == 0]

    ax.bar(['Success', 'Failure'], [len(succ), len(fail)],
           color=['#27ae60', '#e74c3c'], edgecolor='black')
    ax.set_title(site, fontsize=14, fontweight='bold')
    ax.set_ylabel('Count')
    for i, v in enumerate([len(succ), len(fail)]):
        ax.text(i, v + 0.3, str(v), ha='center', fontweight='bold', fontsize=13)

plt.suptitle('Launch Success vs Failure by Site', fontsize=16, fontweight='bold', y=1.02)
plt.tight_layout()
plt.savefig(os.path.join(IMG, "folium_site_success_failure.png"), dpi=150, bbox_inches='tight')
plt.close()
print("Saved folium_site_success_failure.png")

print("All map images generated.")
