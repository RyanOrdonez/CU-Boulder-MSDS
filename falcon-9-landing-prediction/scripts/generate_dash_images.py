"""
Generate static Plotly chart images that mirror the Dash dashboard.
These images will be embedded in the PowerPoint presentation.
"""
import pandas as pd
import plotly.express as px
import os

IMG = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "images")
os.makedirs(IMG, exist_ok=True)

df = pd.read_csv(
    "https://cf-courses-data.s3.us.cloud-object-storage.appdomain.cloud/"
    "IBM-DS0321EN-SkillsNetwork/datasets/spacex_launch_dash.csv"
)

# 1. Pie chart — All Sites total success
fig1 = px.pie(df, values='class', names='Launch Site',
              title='Total Successful Launches by Site')
fig1.update_layout(width=900, height=600)
fig1.write_image(os.path.join(IMG, "dash_pie_all_sites.png"), scale=2)
print("Saved dash_pie_all_sites.png")

# 2. Pie chart — KSC LC-39A success vs failure
ksc = df[df['Launch Site'] == 'KSC LC-39A']
counts = ksc['class'].value_counts().reset_index()
counts.columns = ['class', 'count']
counts['label'] = counts['class'].map({1: 'Success', 0: 'Failure'})
fig2 = px.pie(counts, values='count', names='label',
              title='Success vs Failure — KSC LC-39A')
fig2.update_layout(width=900, height=600)
fig2.write_image(os.path.join(IMG, "dash_pie_ksc.png"), scale=2)
print("Saved dash_pie_ksc.png")

# 3. Scatter — All Sites payload vs outcome
fig3 = px.scatter(df, x='Payload Mass (kg)', y='class',
                  color='Booster Version Category',
                  title='Payload vs Outcome — All Sites',
                  labels={'class': 'Launch Outcome (1=Success)'})
fig3.update_layout(width=1000, height=600)
fig3.write_image(os.path.join(IMG, "dash_scatter_all_sites.png"), scale=2)
print("Saved dash_scatter_all_sites.png")

# 4. Scatter — KSC LC-39A payload vs outcome
fig4 = px.scatter(ksc, x='Payload Mass (kg)', y='class',
                  color='Booster Version Category',
                  title='Payload vs Outcome — KSC LC-39A',
                  labels={'class': 'Launch Outcome (1=Success)'})
fig4.update_layout(width=1000, height=600)
fig4.write_image(os.path.join(IMG, "dash_scatter_ksc.png"), scale=2)
print("Saved dash_scatter_ksc.png")

print("All Dash chart images generated successfully.")
