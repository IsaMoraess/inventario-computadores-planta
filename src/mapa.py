from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from PIL import Image

BASE_DIR = Path(__file__).resolve().parent.parent

PLANTA_PATH = BASE_DIR / "assets" / "planta_real.png"
CSV_PATH = BASE_DIR / "data" / "computadores.csv"
OUTPUT_PATH = BASE_DIR / "output" / "mapa_computadores.html"

planta = Image.open(PLANTA_PATH).convert("RGB")
df = pd.read_csv(CSV_PATH)

total = len(df)
contagem_status = df["status"].value_counts()

ativos = contagem_status.get("Ativo", 0)
manutencao = contagem_status.get("Manutenção", 0)
desligados = contagem_status.get("Desligado", 0)
reservas = contagem_status.get("Reserva", 0)

resumo = (
    f"Total: {total} | "
    f"Ativos: {ativos} | "
    f"Manutenção: {manutencao} | "
    f"Desligados: {desligados} | "
    f"Reservas: {reservas}"
)

fig = px.imshow(planta)

cores = {
    "Ativo": "#2E7D32",
    "Manutenção": "#FB8C00",
    "Desligado": "#D32F2F",
    "Reserva": "#1976D2",
}

for status, grupo in df.groupby("status"):
    fig.add_trace(
        go.Scatter(
            x=grupo["x"],
            y=grupo["y"],
            mode="markers",
            name=status,
            marker=dict(
                size=12,
                color=cores.get(status, "#616161"),
                line=dict(color="white", width=1.5),
            ),
            customdata=grupo[["id", "sala", "Armazenamento", "patrimonio", "status"]],
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Sala: %{customdata[1]}<br>"
                "Armazenamento: %{customdata[2]}<br>"
                "Placa de Vídeo: %{customdata[3]}<br>"
                "Status: %{customdata[4]}"
                "<extra></extra>"
            ),
        )
    )

fig.update_layout(
    title=dict(
	 text=(
	   "Inventário de Computadores - Planta Interativa"
	   f"<br><sup>{resumo}</sup>"
	 ),
	 x=0.02,
	 xanchor="left"
    ),
    width=1200,
    height=800,
    coloraxis_showscale=False,
    paper_bgcolor="white",
    plot_bgcolor="white",
    hoverlabel=dict(
        bgcolor="white",
        bordercolor="#1976D2",
        font_size=13,
        font_family="Arial",
        font_color="black",
        namelength=-1,
    ),
    legend=dict(
        title="Status",
        orientation="v",
        x=1.02,
        y=0.95,
        bgcolor="rgba(255,255,255,0.8)",
    ),
    margin=dict(l=20, r=160, t=70, b=20),
)

fig.update_xaxes(visible=False)
fig.update_yaxes(visible=False)

OUTPUT_PATH.parent.mkdir(exist_ok=True)
fig.write_html(OUTPUT_PATH)

print(f"Mapa gerado com sucesso em: {OUTPUT_PATH}")
