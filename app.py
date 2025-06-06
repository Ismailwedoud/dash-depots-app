import pandas as pd
import dash
from dash import dcc, html, Input, Output, State
import plotly.graph_objects as go
from dash import dash_table
import io
import locale
import base64
from flask import send_file

# Charger les données
df = pd.read_excel("D:/BEA/Contrôle de Gestion/Projet Analyse des agences/Base de données Dépôts 01-2024 au 04-2025 reduite.xlsx", sheet_name="Feuil2")

# Nettoyage
df = df.dropna(subset=["Client", "Compte", "Types de Dépôts", "Année", "Mois", "Dépôts", "Désignation "])
df["Dépôts"] = pd.to_numeric(df["Dépôts"], errors='coerce')
df = df[df["Dépôts"] > 0]
df = df.rename(columns={"Désignation ": "Agence"})


# Initialisation de l'application Dash
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Analyse Dépôts Clients"
server = app.server

exported_table = pd.DataFrame()

def generate_client_month_table(df, agence, type_depot):
    try:
        locale.setlocale(locale.LC_TIME, 'fr_FR.UTF-8')
    except:
        try:
            locale.setlocale(locale.LC_TIME, 'fr_FR')
        except:
            pass

    dff = df[
        (df["Agence"] == agence) &
        (df["Types de Dépôts"] == type_depot) &
        (df["Année"] >= 2024)
    ]

    if dff.empty:
        return pd.DataFrame()

    dff["Date"] = pd.to_datetime(dff["Année"].astype(int).astype(str) + "-" + dff["Mois"].astype(int).astype(str) + "-01")
    dff = dff[dff["Date"] <= "2025-04-30"]
    dff["MoisNom"] = dff["Date"].dt.strftime('%B %Y')

    grouped = dff.groupby(["MoisNom", "Client"])["Dépôts"].sum().reset_index()
    grouped["Rang"] = grouped.groupby("MoisNom")["Dépôts"].rank(ascending=False, method="first").astype(int)

    all_months = pd.date_range("2024-01-01", "2025-04-01", freq='MS').strftime('%B %Y')

    tables = []
    for mois in all_months:
        sub = grouped[grouped["MoisNom"] == mois].copy()
        if sub.empty:
            continue
        sub = sub.set_index("Client")[["Rang", "Dépôts"]]
        sub.columns = pd.MultiIndex.from_product([[mois], ["Rang", "Montant en MRU"]])
        tables.append(sub)

    if not tables:
        return pd.DataFrame()

    result = pd.concat(tables, axis=1)
    result = result.sort_index()
    if ("avril 2025", "Rang") in result.columns:
        result = result.sort_values(by=("avril 2025", "Rang"), na_position="last")

    result.columns = [" - ".join(col) if isinstance(col, tuple) else col for col in result.columns]
    result.reset_index(inplace=True)
    return result

# 🧱 Layout de l'application Dash
app.layout = html.Div([
    html.H2("Classement des clients par dépôts", style={'textAlign': 'center'}),

    html.Div([
        dcc.Dropdown(id='type-depot', options=[{"label": val, "value": val} for val in sorted(df["Types de Dépôts"].unique())], value=df["Types de Dépôts"].iloc[0], style={'width': '50%'}),
        dcc.Dropdown(id='agence', options=[{"label": val, "value": val} for val in sorted(df["Agence"].unique())], value=df["Agence"].iloc[0], style={'width': '70%'}),
        dcc.Dropdown(id='annee', options=[{"label": str(val), "value": val} for val in sorted(df["Année"].unique())], value=df["Année"].iloc[0], style={'width': '40%'}),
        dcc.Dropdown(id='mois', options=[{"label": str(val), "value": val} for val in sorted(df["Mois"].unique())], value=df["Mois"].iloc[0], style={'width': '30%'})
    ], style={'display': 'flex', 'gap': '15px', 'justifyContent': 'center'}),

    html.Br(),

    html.Div([
        html.Label("Nombre de clients à afficher (Top N)"),
        dcc.Slider(
            id='top-n-slider',
            min=5,
            max=100,
            step=5,
            value=25,
            marks={i: str(i) for i in range(5, 200, 10)},
            tooltip={"placement": "bottom", "always_visible": True},
            updatemode='drag'
        )
    ], style={'margin': '0 10% 20px'}),

    dcc.Graph(id='graph-depots'),
    html.H4("Statistiques des dépôts", style={'textAlign': 'center'}),
    html.Div(id='table-depots'),
    html.H4("Rang et Montant par Client et Mois", style={'textAlign': 'center'}),
    html.Div(id='rang-montant-table'),
    html.Br(),
    html.Div([
        html.Button("Exporter en Excel", id="export-btn"),
        dcc.Download(id="download-excel")
    ], style={"textAlign": "center"})
])

# 📤 Callback pour exporter le tableau sous Excel
@app.callback(
    Output("download-excel", "data"),
    Input("export-btn", "n_clicks"),
    State("type-depot", "value"),
    State("agence", "value"),
    prevent_initial_call=True
)
def export_to_excel(n_clicks, type_depot, agence):
    if n_clicks:
        export_data = generate_client_month_table(df, agence, type_depot)
        if export_data.empty:
            return dash.no_update
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            export_data.to_excel(writer, index=False, sheet_name="Rang_Montant")
        output.seek(0)
        return dcc.send_bytes(output.read(), filename="rang_montant_clients.xlsx")

# ✅ Callback principal pour afficher le graphique et les tables
@app.callback(
    Output('graph-depots', 'figure'),
    Output('table-depots', 'children'),
    Output('rang-montant-table', 'children'),
    Input('type-depot', 'value'),
    Input('agence', 'value'),
    Input('annee', 'value'),
    Input('mois', 'value'),
    Input('top-n-slider', 'value')
)
def update_graph_and_tables(type_depot, agence, annee, mois, top_n):
    dff = df[
        (df["Types de Dépôts"] == type_depot) &
        (df["Agence"] == agence) &
        (df["Année"] == annee) &
        (df["Mois"] == mois)
    ]

    if dff.empty:
        fig = go.Figure().update_layout(title="Aucune donnée disponible")
        return fig, html.Div("Pas de données disponibles"), html.Div()

    dff_grouped = dff.groupby(["Client", "Compte"])["Dépôts"].sum().reset_index()
    dff_grouped["Rang"] = dff_grouped["Dépôts"].rank(method="first", ascending=False)
    dff_grouped = dff_grouped.sort_values("Dépôts", ascending=False).head(top_n)

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=dff_grouped["Rang"],
        y=dff_grouped["Dépôts"],
        text=dff_grouped["Client"],
        textposition="auto",
        hovertemplate="<b>Client:</b> %{text}<br><b>Montant:</b> %{y:,.0f} MRU<br><b>Rang:</b> %{x}",
        name="Clients"
    ))

    fig.update_layout(
        title=f"{type_depot} - {agence} ({mois}/{annee})",
        xaxis_title="Rang",
        yaxis_title="Montant déposé (MRU)",
        height=500,
        plot_bgcolor="white"
    )

    montants = dff_grouped["Dépôts"]
    seuil_80 = montants.quantile(0.80)
    sup_80 = dff_grouped[montants >= seuil_80]
    inf_80 = dff_grouped[montants < seuil_80]
    total = montants.sum()
    exposition = (sup_80["Dépôts"].sum() / total * 100) if total > 0 else 0

    stats = pd.DataFrame({
        "Indicateur": [
            "Nombre total de clients",
            "Montant total (MRU)",
            "Nombre de Clients ≥ 80e percentile",
            "Montant ≥ 80e percentile",
            "Nombre Clients < 80e percentile",
            "Montant < 80e percentile",
            "Exposition (≥ 80e percentile)"
        ],
        "Valeur": [
            len(dff_grouped),
            f"{total:,.0f}".replace(",", " "),
            len(sup_80),
            f"{sup_80['Dépôts'].sum():,.0f}".replace(",", " "),
            len(inf_80),
            f"{inf_80['Dépôts'].sum():,.0f}".replace(",", " "),
            f"{exposition:.2f}%"
        ]
    })

    table_stats = dash_table.DataTable(
        columns=[{"name": i, "id": i} for i in stats.columns],
        data=stats.to_dict("records"),
        style_cell={"textAlign": "center"},
        style_header={"fontWeight": "bold", "backgroundColor": "#ddd"},
        style_table={"margin": "auto", "width": "60%"}
    )

    table_rang = generate_client_month_table(df, agence, type_depot)
    if table_rang.empty:
        return fig, table_stats, html.Div("Pas de tableau à afficher")

    table_html = dash_table.DataTable(
        columns=[{"name": col, "id": col} for col in table_rang.columns],
        data=table_rang.to_dict("records"),
        style_cell={'textAlign': 'center', 'padding': '5px'},
        style_header={'backgroundColor': '#eee', 'fontWeight': 'bold'},
        style_table={'overflowX': 'auto'}
    )

    return fig, table_stats, table_html

# 🚀 Lancement de l'app
if __name__ == '__main__':
    app.run(debug=True, port=8060)