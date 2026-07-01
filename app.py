from pathlib import Path
from datetime import datetime
from io import BytesIO
import base64
import os
import re
from textwrap import shorten
from unicodedata import normalize
from urllib.parse import quote
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from PIL import Image, ImageDraw
from reportlab.graphics.charts.barcharts import HorizontalBarChart, VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics.shapes import Drawing, String
from reportlab.lib import colors as rl_colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfgen import canvas
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from streamlit_image_coordinates import streamlit_image_coordinates
from streamlit_plotly_events import plotly_events

try:
    from dotenv import load_dotenv
except ImportError:
    load_dotenv = None

try:
    import qrcode
except ImportError:
    qrcode = None

try:
    from sqlalchemy import create_engine, text
except ImportError:
    create_engine = None
    text = None

BASE_DIR = Path(__file__).resolve().parent

if load_dotenv is not None:
    load_dotenv(BASE_DIR / ".env")

# PLANTA_PATH = BASE_DIR / "assets" / "planta.png"

# Planta real reconstruída pelo contorno do prédio.
# Usar quando for trocar a base do sistema:
PLANTA_PATH = BASE_DIR / "assets" / "planta_real.png"
LOGO_PATH = BASE_DIR / "assets" / "logo_jr.png"

CSV_PATH = BASE_DIR / "data" / "computadores.csv"
BACKUP_DIR = BASE_DIR / "data" / "backups"
MOVIMENTACOES_PATH = BASE_DIR / "data" / "movimentacoes.csv"

CAMPOS_CSV = [
    "id",
    "sala",
    "x",
    "y",
    "status",
    "armazenamento",
    "placa_video",
    "usuario",
    "sistema",
    "ram",
    "processador",
    "observacoes",
]
CAMPOS_MOVIMENTACAO = [
    "id",
    "data_hora",
    "computador_id",
    "campo",
    "valor_anterior",
    "valor_novo",
    "acao",
]
CAMPOS_MOVIMENTACAO_MONITORADOS = [
    "status",
    "armazenamento",
    "placa_video",
    "sistema",
    "ram",
    "processador",
    "observacoes",
    "sala",
    "usuario",
    "x",
    "y",
]
SISTEMAS_ANTIGOS = ["Windows 7", "Windows 8", "Windows 8.1"]
PALAVRAS_CRITICAS = ["urgente", "critico", "crítico", "falha", "troca", "fonte", "defeito"]

st.set_page_config(page_title="Gestão de Ativos de TI", layout="wide")

st.markdown(
    """
    <style>
    section[data-testid="stSidebar"] .stMarkdown p {
        margin-bottom: 0.35rem;
    }
    .asset-card {
        border: 1px solid rgba(148, 163, 184, 0.28);
        border-radius: 8px;
        padding: 12px 12px 6px 12px;
        margin: 8px 0 10px 0;
        background: rgba(255, 255, 255, 0.04);
    }
    .asset-row {
        display: grid;
        grid-template-columns: 92px 1fr;
        gap: 8px;
        padding: 5px 0;
        border-bottom: 1px solid rgba(148, 163, 184, 0.14);
        font-size: 0.9rem;
    }
    .asset-row:last-child {
        border-bottom: 0;
    }
    .asset-label {
        color: rgba(226, 232, 240, 0.72);
        font-weight: 600;
    }
    .asset-value {
        color: rgba(248, 250, 252, 0.96);
        overflow-wrap: anywhere;
    }
    .dash-card-grid {
        display: grid;
        grid-template-columns: repeat(4, minmax(0, 1fr));
        gap: 12px;
        margin: 10px 0 16px 0;
    }
    .dash-card {
        border: 1px solid rgba(148, 163, 184, 0.22);
        border-radius: 8px;
        padding: 14px 16px;
        background: linear-gradient(180deg, rgba(15, 23, 42, 0.88), rgba(15, 23, 42, 0.62));
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.18);
    }
    .dash-card-label {
        color: #94A3B8;
        font-size: 0.78rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0;
        margin-bottom: 8px;
    }
    .dash-card-value {
        color: #F8FAFC;
        font-size: 1.7rem;
        line-height: 1.1;
        font-weight: 800;
    }
    .executive-summary {
        border: 1px solid rgba(59, 130, 246, 0.22);
        border-left: 4px solid #3B82F6;
        border-radius: 8px;
        padding: 14px 16px;
        margin: 4px 0 18px 0;
        background: rgba(30, 41, 59, 0.7);
        color: #E2E8F0;
        font-size: 0.98rem;
        line-height: 1.45;
    }
    @media (max-width: 1200px) {
        .dash-card-grid {
            grid-template-columns: repeat(2, minmax(0, 1fr));
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def dataframe_para_xlsx(dados):
    def coluna_excel(indice):
        letras = ""

        while indice:
            indice, resto = divmod(indice - 1, 26)
            letras = chr(65 + resto) + letras

        return letras

    linhas = [dados.columns.tolist()] + dados.astype(str).values.tolist()
    planilha = []

    for numero_linha, linha in enumerate(linhas, start=1):
        celulas = []

        for numero_coluna, valor in enumerate(linha, start=1):
            referencia = f"{coluna_excel(numero_coluna)}{numero_linha}"
            celulas.append(
                f'<c r="{referencia}" t="inlineStr">'
                f"<is><t>{escape(valor)}</t></is>"
                "</c>"
            )

        planilha.append(f'<row r="{numero_linha}">{"".join(celulas)}</row>')

    ultima_coluna = coluna_excel(len(dados.columns))
    ultima_linha = len(linhas)
    arquivo = BytesIO()

    with ZipFile(arquivo, "w", ZIP_DEFLATED) as xlsx:
        xlsx.writestr(
            "[Content_Types].xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>""",
        )
        xlsx.writestr(
            "_rels/.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
        )
        xlsx.writestr(
            "xl/workbook.xml",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Inventario" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
        )
        xlsx.writestr(
            "xl/_rels/workbook.xml.rels",
            """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>""",
        )
        xlsx.writestr(
            "xl/worksheets/sheet1.xml",
            f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="A1:{ultima_coluna}{ultima_linha}"/>
<sheetData>{"".join(planilha)}</sheetData>
</worksheet>""",
        )

    return arquivo.getvalue()


def cor_hex_para_rgb(cor):
    cor = cor.lstrip("#")
    return tuple(int(cor[indice : indice + 2], 16) for indice in (0, 2, 4))


def agora_texto():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def normalizar_dataframe(dados):
    dados = dados.copy()

    colunas_armazenamento = {"ip", "Armazenamento", "armazenamento"}
    if "processamento" in dados.columns and not colunas_armazenamento.intersection(dados.columns):
        if "placa de vídeo" in dados.columns:
            dados["armazenamento"] = dados["placa de vídeo"]
        elif "placa de vÃ­deo" in dados.columns:
            dados["armazenamento"] = dados["placa de vÃ­deo"]

        dados["placa_video"] = dados["processamento"]

    aliases = {
        "ip": "armazenamento",
        "Armazenamento": "armazenamento",
        "patrimonio": "placa_video",
        "placa de vídeo": "placa_video",
        "placa de vÃ­deo": "placa_video",
        "processamento": "placa_video",
    }

    for coluna_antiga, coluna_nova in aliases.items():
        if coluna_antiga in dados.columns:
            if coluna_nova not in dados.columns:
                dados[coluna_nova] = ""

            mascara_vazia = dados[coluna_nova].fillna("").astype(str).str.strip() == ""
            dados.loc[mascara_vazia, coluna_nova] = dados.loc[mascara_vazia, coluna_antiga]

    for coluna in CAMPOS_CSV:
        if coluna not in dados.columns:
            dados[coluna] = ""

    dados = dados.reindex(columns=CAMPOS_CSV).fillna("")

    for coluna in ["x", "y"]:
        dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce").fillna(0).astype(int)

    return dados


def obter_url_banco():
    try:
        url_secrets = st.secrets.get("SUPABASE_DB_URL", "")
    except Exception:
        url_secrets = ""

    return (url_secrets or os.getenv("SUPABASE_DB_URL", "")).strip()


def normalizar_url_postgres(url):
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif url.startswith("postgresql://"):
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)

    if "sslmode=" not in url and "localhost" not in url and "127.0.0.1" not in url:
        separador = "&" if "?" in url else "?"
        url = f"{url}{separador}sslmode=require"

    return url


@st.cache_resource
def conectar_banco():
    if create_engine is None:
        raise RuntimeError(
            "Dependência ausente: instale sqlalchemy e psycopg2-binary para conectar ao Supabase."
        )

    url = obter_url_banco()

    if not url:
        raise RuntimeError(
            "SUPABASE_DB_URL não configurada. Defina essa variável no ambiente ou em .streamlit/secrets.toml."
        )

    return create_engine(normalizar_url_postgres(url), pool_pre_ping=True)


def inicializar_banco():
    engine = conectar_banco()

    with engine.begin() as conexao:
        conexao.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS computadores (
                    id TEXT PRIMARY KEY,
                    sala TEXT,
                    x INTEGER,
                    y INTEGER,
                    status TEXT,
                    armazenamento TEXT,
                    placa_video TEXT,
                    usuario TEXT,
                    sistema TEXT,
                    ram TEXT,
                    processador TEXT,
                    observacoes TEXT
                )
                """
            )
        )
        conexao.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS movimentacoes (
                    id BIGSERIAL PRIMARY KEY,
                    data_hora TIMESTAMP,
                    computador_id TEXT,
                    campo TEXT,
                    valor_anterior TEXT,
                    valor_novo TEXT,
                    acao TEXT
                )
                """
            )
        )


def importar_csv_inicial():
    engine = conectar_banco()

    with engine.begin() as conexao:
        total_computadores = conexao.execute(text("SELECT COUNT(*) FROM computadores")).scalar() or 0
        total_movimentacoes = conexao.execute(text("SELECT COUNT(*) FROM movimentacoes")).scalar() or 0

    importados = 0

    if total_computadores == 0 and CSV_PATH.exists():
        dados = normalizar_dataframe(pd.read_csv(CSV_PATH).fillna(""))
        registros = dados.to_dict("records")
        importar_computadores(registros, ignorar_duplicados=True)
        importados = len(registros)

    if total_movimentacoes == 0 and MOVIMENTACOES_PATH.exists():
        historico = pd.read_csv(MOVIMENTACOES_PATH).fillna("")
        registros_historico = []

        for _, linha in historico.iterrows():
            computador_id = linha.get("computador_id", linha.get("id", ""))
            registros_historico.append(
                {
                    "data_hora": str(linha.get("data_hora", "")) or agora_texto(),
                    "computador_id": str(computador_id),
                    "campo": str(linha.get("campo", "")),
                    "valor_anterior": str(linha.get("valor_anterior", "")),
                    "valor_novo": str(linha.get("valor_novo", "")),
                    "acao": str(linha.get("acao", "")),
                }
            )

        if registros_historico:
            with engine.begin() as conexao:
                conexao.execute(
                    text(
                        """
                        INSERT INTO movimentacoes (
                            data_hora,
                            computador_id,
                            campo,
                            valor_anterior,
                            valor_novo,
                            acao
                        )
                        VALUES (
                            :data_hora,
                            :computador_id,
                            :campo,
                            :valor_anterior,
                            :valor_novo,
                            :acao
                        )
                        """
                    ),
                    registros_historico,
                )

    return importados


def carregar_computadores():
    engine = conectar_banco()

    with engine.connect() as conexao:
        dados = pd.read_sql_query(
            text(f"SELECT {', '.join(CAMPOS_CSV)} FROM computadores ORDER BY id"),
            conexao,
        )

    return normalizar_dataframe(dados)


def carregar_movimentacoes():
    engine = conectar_banco()

    with engine.connect() as conexao:
        return pd.read_sql_query(
            text(
                """
                SELECT id, data_hora, computador_id, campo, valor_anterior, valor_novo, acao
                FROM movimentacoes
                ORDER BY id
                """
            ),
            conexao,
        ).fillna("")


def salvar_computador(registro):
    registro = normalizar_dataframe(pd.DataFrame([registro])).iloc[0].to_dict()
    engine = conectar_banco()
    placeholders = ", ".join([f":{campo}" for campo in CAMPOS_CSV])

    with engine.begin() as conexao:
        conexao.execute(
            text(
                f"""
                INSERT INTO computadores ({", ".join(CAMPOS_CSV)})
                VALUES ({placeholders})
                """
            ),
            registro,
        )


def importar_computadores(registros, ignorar_duplicados=False):
    if not registros:
        return

    engine = conectar_banco()
    placeholders = ", ".join([f":{campo}" for campo in CAMPOS_CSV])
    conflito = " ON CONFLICT (id) DO NOTHING" if ignorar_duplicados else ""

    with engine.begin() as conexao:
        conexao.execute(
            text(
                f"""
                INSERT INTO computadores ({", ".join(CAMPOS_CSV)})
                VALUES ({placeholders})
                {conflito}
                """
            ),
            normalizar_dataframe(pd.DataFrame(registros)).to_dict("records"),
        )


def atualizar_computador(computador_id, registro, acao="Edição"):
    dados_atuais = carregar_computadores()
    computador_atual = dados_atuais[dados_atuais["id"] == computador_id]

    if computador_atual.empty:
        return False

    registro_antigo = computador_atual.iloc[0].to_dict()
    registro_novo = normalizar_dataframe(pd.DataFrame([registro])).iloc[0].to_dict()
    engine = conectar_banco()

    atribuicoes = ", ".join([f"{campo} = :{campo}" for campo in CAMPOS_CSV])
    parametros = {**registro_novo, "id_original": computador_id}

    with engine.begin() as conexao:
        conexao.execute(
            text(f"UPDATE computadores SET {atribuicoes} WHERE id = :id_original"),
            parametros,
        )

    registrar_movimentacoes(registro_antigo, registro_novo, acao)
    return True


def excluir_computador(computador_id):
    engine = conectar_banco()

    with engine.begin() as conexao:
        conexao.execute(
            text("DELETE FROM computadores WHERE id = :computador_id"),
            {"computador_id": computador_id},
        )


def valor_limpo(valor):
    return str(valor).strip()


def texto_normalizado(valor):
    return normalize("NFKD", valor_limpo(valor)).encode("ascii", "ignore").decode("ascii").lower()


def valor_nao_informado(valor):
    return texto_normalizado(valor) in {"", "a coletar", "nao informado", "sem informacao"}


def pouca_ram(valor):
    if valor_nao_informado(valor):
        return False

    texto = texto_normalizado(valor).replace(",", ".")
    encontrado = re.search(r"\d+(?:\.\d+)?", texto)

    if not encontrado:
        return False

    quantidade = float(encontrado.group())

    if "mb" in texto:
        quantidade = quantidade / 1024

    return quantidade <= 4


def ram_em_gb(valor):
    texto = texto_normalizado(valor).replace(",", ".")
    encontrado = re.search(r"\d+(?:\.\d+)?", texto)

    if not encontrado:
        return None

    quantidade = float(encontrado.group())

    if "mb" in texto:
        quantidade = quantidade / 1024

    return quantidade


def sistema_antigo(valor):
    if valor_nao_informado(valor):
        return False

    sistema = texto_normalizado(valor)
    return sistema in {texto_normalizado(item) for item in SISTEMAS_ANTIGOS}


def valor_preenchido(valor):
    return texto_normalizado(valor) not in {
        "",
        "a coletar",
        "nao informado",
        "sem informacao",
        "ninguem",
        "sem usuario",
        "vazio",
    }


def campo_pendente(valor):
    return not valor_preenchido(valor)


def padronizar_ram(valor):
    if campo_pendente(valor):
        return "A coletar"

    quantidade = ram_em_gb(valor)

    if quantidade is None:
        return "A coletar"

    if abs(quantidade - round(quantidade)) < 0.05:
        quantidade = int(round(quantidade))

    if quantidade in [4, 6, 8, 16, 32]:
        return f"{quantidade} GB"

    return "Outros"


def padronizar_sistema(valor):
    if campo_pendente(valor):
        return "A coletar"

    sistema = texto_normalizado(valor)

    if sistema == "windows 10":
        return "Windows 10"

    if sistema == "windows 11":
        return "Windows 11"

    return "Outros"


def campos_pendentes(item):
    campos = ["armazenamento", "placa_video", "sistema", "ram", "processador", "usuario"]
    return [campo for campo in campos if campo_pendente(item.get(campo, ""))]


def tabela_pendencias(dados):
    linhas = []

    for _, item in dados.iterrows():
        pendentes = campos_pendentes(item)

        if pendentes:
            linhas.append(
                {
                    "ID": item["id"],
                    "Sala": item["sala"],
                    "Usuário": item["usuario"],
                    "Campos pendentes": ", ".join(pendentes),
                }
            )

    return pd.DataFrame(linhas, columns=["ID", "Sala", "Usuário", "Campos pendentes"])


def percentual_inventario_completo(dados):
    campos = ["sala", "status", "armazenamento", "placa_video", "usuario", "sistema", "ram", "processador"]

    if dados.empty:
        return 0

    total_campos = len(dados) * len(campos)
    preenchidos = 0

    for campo in campos:
        preenchidos += dados[campo].map(valor_preenchido).sum()

    return round((preenchidos / total_campos) * 100) if total_campos else 0


def contar_alertas_por_tipo(dados):
    contagem = {
        "Manutenção": 0,
        "Desligado": 0,
        "Pouca RAM": 0,
        "Sem usuário": 0,
        "Sistema antigo": 0,
        "Observação crítica": 0,
    }

    for _, item in dados.iterrows():
        motivos = motivos_alerta(item)

        for motivo in motivos:
            if motivo in contagem:
                contagem[motivo] += 1

    return pd.DataFrame(
        [{"tipo": tipo, "total": total} for tipo, total in contagem.items()]
    )


def render_dashboard_card(titulo, valor, cor="#3B82F6"):
    return (
        f"<div class='dash-card' style='border-top: 3px solid {cor};'>"
        f"<div class='dash-card-label'>{escape(str(titulo))}</div>"
        f"<div class='dash-card-value'>{escape(str(valor))}</div>"
        "</div>"
    )


def url_computador(computador_id):
    base_url = st.session_state.get("app_base_url", "http://localhost:8503")
    return f"{base_url.rstrip('/')}?pc={quote(str(computador_id))}"


def qr_code_png(conteudo):
    if qrcode is None:
        return None

    qr = qrcode.QRCode(version=1, box_size=8, border=3)
    qr.add_data(conteudo)
    qr.make(fit=True)
    imagem_qr = qr.make_image(fill_color="black", back_color="white").convert("RGB")
    arquivo = BytesIO()
    imagem_qr.save(arquivo, format="PNG")
    return arquivo.getvalue()


def computador_clicado(selected_points, ids_por_trace, dados):
    if not selected_points:
        return None

    for ponto in selected_points:
        numero_trace = ponto.get("curveNumber")
        numero_ponto = ponto.get("pointNumber")

        if numero_trace is not None:
            numero_trace = int(numero_trace)

        if numero_ponto is not None:
            numero_ponto = int(numero_ponto)

        ids_do_trace = ids_por_trace.get(numero_trace, [])

        if numero_ponto is not None and numero_ponto < len(ids_do_trace):
            return ids_do_trace[numero_ponto]

    ponto = selected_points[0]
    ponto_x = ponto.get("x")
    ponto_y = ponto.get("y")
    computador_encontrado = dados[(dados["x"] == ponto_x) & (dados["y"] == ponto_y)]

    if not computador_encontrado.empty:
        return computador_encontrado.iloc[0]["id"]

    return None


def coordenadas_clicadas(selected_points):
    if not selected_points:
        return None

    ponto = selected_points[0]

    if ponto.get("x") is None or ponto.get("y") is None:
        return None

    return int(round(float(ponto["x"]))), int(round(float(ponto["y"])))


def salvar_posicao_computador(dados, computador_id, novo_x, novo_y):
    mascara_computador = dados["id"] == computador_id

    if not mascara_computador.any():
        return dados, False

    registro_novo = dados.loc[mascara_computador].iloc[0].to_dict()
    registro_novo["x"] = int(novo_x)
    registro_novo["y"] = int(novo_y)

    salvou = atualizar_computador(computador_id, registro_novo, "Posicionamento")
    return carregar_computadores(), salvou


def texto_pdf(valor):
    texto = normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return texto


def validar_registro(registro, dados_base, id_original=None):
    erros = []
    computador_id = valor_limpo(registro.get("id", ""))
    status = valor_limpo(registro.get("status", ""))

    if not computador_id:
        erros.append("Informe o nome do computador.")

    if status not in cores:
        erros.append("Informe um status válido.")

    for campo in ["x", "y"]:
        try:
            int(registro.get(campo, 0))
        except (TypeError, ValueError):
            erros.append(f"Informe uma coordenada {campo.upper()} válida.")

    x = int(registro.get("x", 0) or 0)
    y = int(registro.get("y", 0) or 0)

    if not 0 <= x <= largura_planta:
        erros.append(f"A coordenada X deve ficar entre 0 e {largura_planta}.")

    if not 0 <= y <= altura_planta:
        erros.append(f"A coordenada Y deve ficar entre 0 e {altura_planta}.")

    comparacao = dados_base.copy()

    if id_original is not None:
        comparacao = comparacao[comparacao["id"] != id_original]

    if computador_id and computador_id in comparacao["id"].astype(str).tolist():
        erros.append("Já existe um computador com esse nome.")

    return erros


def registrar_movimentacoes(registro_antigo, registro_novo, acao):
    linhas = []

    for campo in CAMPOS_MOVIMENTACAO_MONITORADOS:
        valor_anterior = str(registro_antigo.get(campo, ""))
        valor_novo = str(registro_novo.get(campo, ""))

        if valor_anterior != valor_novo:
            linhas.append(
                {
                    "data_hora": agora_texto(),
                    "computador_id": registro_novo.get("id", registro_antigo.get("id", "")),
                    "campo": campo,
                    "valor_anterior": valor_anterior,
                    "valor_novo": valor_novo,
                    "acao": acao,
                }
            )

    if not linhas:
        return

    engine = conectar_banco()

    with engine.begin() as conexao:
        conexao.execute(
            text(
                """
                INSERT INTO movimentacoes (
                    data_hora,
                    computador_id,
                    campo,
                    valor_anterior,
                    valor_novo,
                    acao
                )
                VALUES (
                    :data_hora,
                    :computador_id,
                    :campo,
                    :valor_anterior,
                    :valor_novo,
                    :acao
                )
                """
            ),
            linhas,
        )


def motivos_alerta(item):
    motivos = []
    status = str(item.get("status", ""))
    usuario = valor_limpo(item.get("usuario", ""))
    ram = valor_limpo(item.get("ram", "")).lower()
    sistema = valor_limpo(item.get("sistema", ""))
    observacoes = valor_limpo(item.get("observacoes", ""))

    if status in ["Manutenção", "Desligado"]:
        motivos.append(status)

    if not usuario or usuario.lower() == "sem usuário":
        motivos.append("Sem usuário")

    if pouca_ram(ram):
        motivos.append("Pouca RAM")

    if sistema_antigo(sistema):
        motivos.append("Sistema antigo")

    if any(palavra in observacoes.lower() for palavra in PALAVRAS_CRITICAS):
        motivos.append("Observação crítica")

    return motivos


def dados_com_alertas(dados):
    dados = dados.copy()
    dados["motivos_alerta"] = dados.apply(motivos_alerta, axis=1)
    return dados[dados["motivos_alerta"].map(bool)]


class RodapeCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._paginas_salvas = []

    def showPage(self):
        self._paginas_salvas.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_paginas = len(self._paginas_salvas)

        for pagina in self._paginas_salvas:
            self.__dict__.update(pagina)
            self.desenhar_rodape(total_paginas)
            super().showPage()

        super().save()

    def desenhar_rodape(self, total_paginas):
        largura, _ = A4
        texto_data = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.setStrokeColor(rl_colors.HexColor("#CBD5E1"))
        self.line(1.6 * cm, 1.25 * cm, largura - 1.6 * cm, 1.25 * cm)
        self.setFillColor(rl_colors.HexColor("#475569"))
        self.setFont("Helvetica", 8)
        self.drawString(1.6 * cm, 0.8 * cm, "Sistema de Inventário de Computadores")
        self.drawCentredString(largura / 2, 0.8 * cm, f"Página {self._pageNumber} de {total_paginas}")
        self.drawRightString(largura - 1.6 * cm, 0.8 * cm, texto_data)


def paragrafo_pdf(texto, estilo):
    return Paragraph(escape(str(texto)), estilo)


def imagem_logo_relatorio():
    return LOGO_PATH if LOGO_PATH.exists() else None


def logo_html_dashboard():
    if not LOGO_PATH.exists():
        return ""

    conteudo = base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
    return (
        '<img src="data:image/png;base64,'
        f'{conteudo}" '
        'style="width: 260px; max-width: 100%; margin-bottom: 16px;" '
        'alt="Logo JR">'
    )


def cor_reportlab(hex_cor):
    return rl_colors.HexColor(hex_cor)


def card_pdf(titulo, valor, cor, estilos):
    tabela = Table(
        [
            [paragrafo_pdf(titulo, estilos["CardTitulo"])],
            [paragrafo_pdf(valor, estilos["CardValor"])],
        ],
        colWidths=[5.2 * cm],
        rowHeights=[0.75 * cm, 1.05 * cm],
    )
    tabela.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), rl_colors.white),
                ("BOX", (0, 0), (-1, -1), 0.8, cor_reportlab("#CBD5E1")),
                ("LINEBEFORE", (0, 0), (0, -1), 4, cor),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return tabela


def grafico_barras_status_pdf(dados, paleta):
    contagem = dados["status"].value_counts()
    labels = contagem.index.astype(str).tolist()
    valores = [int(valor) for valor in contagem.tolist()]
    desenho = Drawing(460, 230)
    grafico = VerticalBarChart()
    grafico.x = 40
    grafico.y = 35
    grafico.height = 145
    grafico.width = 370
    grafico.data = [valores or [0]]
    grafico.categoryAxis.categoryNames = labels or ["Sem dados"]
    grafico.categoryAxis.labels.boxAnchor = "ne"
    grafico.categoryAxis.labels.angle = 30
    grafico.valueAxis.valueMin = 0
    grafico.valueAxis.valueMax = max(valores or [1]) + 1
    grafico.valueAxis.valueStep = max(1, round(grafico.valueAxis.valueMax / 4))
    grafico.bars[0].fillColor = cor_reportlab("#1976D2")
    desenho.add(String(40, 205, "Computadores por status", fontSize=11, fillColor=cor_reportlab("#0F172A")))
    desenho.add(grafico)
    return desenho


def grafico_barras_sala_pdf(dados):
    contagem = dados["sala"].value_counts().head(8)
    labels = contagem.index.astype(str).tolist()
    valores = [int(valor) for valor in contagem.tolist()]
    desenho = Drawing(460, 240)
    grafico = HorizontalBarChart()
    grafico.x = 130
    grafico.y = 35
    grafico.height = 155
    grafico.width = 280
    grafico.data = [valores or [0]]
    grafico.categoryAxis.categoryNames = labels or ["Sem dados"]
    grafico.categoryAxis.labels.fontSize = 7
    grafico.valueAxis.valueMin = 0
    grafico.valueAxis.valueMax = max(valores or [1]) + 1
    grafico.valueAxis.valueStep = max(1, round(grafico.valueAxis.valueMax / 4))
    grafico.bars[0].fillColor = cor_reportlab("#2563EB")
    desenho.add(String(40, 215, "Top salas por quantidade", fontSize=11, fillColor=cor_reportlab("#0F172A")))
    desenho.add(grafico)
    return desenho


def grafico_windows_pdf(dados):
    windows_10 = int((dados["sistema"].astype(str).str.lower() == "windows 10").sum())
    windows_11 = int((dados["sistema"].astype(str).str.lower() == "windows 11").sum())
    desenho = Drawing(460, 220)
    pizza = Pie()
    pizza.x = 135
    pizza.y = 45
    pizza.width = 130
    pizza.height = 130
    pizza.data = [windows_10, windows_11] if windows_10 or windows_11 else [1]
    pizza.labels = ["Windows 10", "Windows 11"] if windows_10 or windows_11 else ["Sem dados"]
    pizza.slices[0].fillColor = cor_reportlab("#60A5FA")
    if len(pizza.data) > 1:
        pizza.slices[1].fillColor = cor_reportlab("#22C55E")
    desenho.add(String(40, 195, "Windows 10 x Windows 11", fontSize=11, fillColor=cor_reportlab("#0F172A")))
    desenho.add(pizza)
    return desenho


def tabela_pdf(cabecalhos, linhas, col_widths, estilos, fonte=7):
    dados_tabela = [[paragrafo_pdf(cabecalho, estilos["TabelaCabecalho"]) for cabecalho in cabecalhos]]

    for linha in linhas:
        dados_tabela.append([paragrafo_pdf(valor, estilos["TabelaCelula"]) for valor in linha])

    tabela = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
    estilos_tabela = [
        ("BACKGROUND", (0, 0), (-1, 0), cor_reportlab("#0F172A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), rl_colors.white),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, cor_reportlab("#CBD5E1")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), fonte),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]

    for indice in range(1, len(dados_tabela)):
        cor_linha = "#F8FAFC" if indice % 2 == 0 else "#FFFFFF"
        estilos_tabela.append(("BACKGROUND", (0, indice), (-1, indice), cor_reportlab(cor_linha)))

    tabela.setStyle(TableStyle(estilos_tabela))
    return tabela


def relatorio_pdf(imagem, dados, alertas, paleta, computador_selecionado):
    arquivo = BytesIO()
    data_geracao = agora_texto()
    doc = SimpleDocTemplate(
        arquivo,
        pagesize=A4,
        rightMargin=1.6 * cm,
        leftMargin=1.6 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.8 * cm,
        title="Relatório de Inventário de Computadores",
    )
    estilos_base = getSampleStyleSheet()
    estilos = {
        "TituloCapa": ParagraphStyle(
            "TituloCapa",
            parent=estilos_base["Title"],
            fontName="Helvetica-Bold",
            fontSize=28,
            leading=34,
            alignment=TA_CENTER,
            textColor=cor_reportlab("#0F172A"),
            spaceAfter=14,
        ),
    "Subtitulo": ParagraphStyle(
        "Subtitulo",
        parent=estilos_base["Heading2"],
            fontName="Helvetica",
            fontSize=15,
            leading=20,
            alignment=TA_CENTER,
            textColor=cor_reportlab("#334155"),
        spaceAfter=24,
    ),
        "MarcaCapa": ParagraphStyle(
            "MarcaCapa",
            parent=estilos_base["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=24,
            alignment=TA_CENTER,
            textColor=cor_reportlab("#0F172A"),
            spaceAfter=10,
        ),
        "Secao": ParagraphStyle(
            "Secao",
            parent=estilos_base["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=24,
            textColor=cor_reportlab("#0F172A"),
            spaceAfter=14,
        ),
        "Texto": ParagraphStyle(
            "Texto",
            parent=estilos_base["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=cor_reportlab("#334155"),
            spaceAfter=8,
        ),
        "CardTitulo": ParagraphStyle(
            "CardTitulo",
            parent=estilos_base["BodyText"],
            fontSize=8,
            leading=10,
            textColor=cor_reportlab("#64748B"),
        ),
        "CardValor": ParagraphStyle(
            "CardValor",
            parent=estilos_base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=cor_reportlab("#0F172A"),
        ),
        "TabelaCabecalho": ParagraphStyle(
            "TabelaCabecalho",
            parent=estilos_base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=7.5,
            leading=9,
            alignment=TA_CENTER,
            textColor=rl_colors.white,
        ),
        "TabelaCelula": ParagraphStyle(
            "TabelaCelula",
            parent=estilos_base["BodyText"],
            fontSize=7,
            leading=9,
            textColor=cor_reportlab("#0F172A"),
        ),
        "AlertaTitulo": ParagraphStyle(
            "AlertaTitulo",
            parent=estilos_base["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=cor_reportlab("#0F172A"),
        ),
    }

    historia = []
    total = len(dados)
    total_ativos = int((dados["status"] == "Ativo").sum())
    total_manutencao = int((dados["status"] == "Manutenção").sum())
    total_desligados = int((dados["status"] == "Desligado").sum())
    total_reserva = int((dados["status"] == "Reserva").sum())
    total_alertas = len(alertas)
    total_salas = dados["sala"].replace("", "Não informado").nunique()

    logo = imagem_logo_relatorio()
    if logo:
        logo_pdf = RLImage(str(logo), width=4 * cm, height=4 * cm, kind="proportional")
        logo_pdf.hAlign = "CENTER"
        historia.append(logo_pdf)
        historia.append(Spacer(1, 0.7 * cm))
    else:
        historia.append(Spacer(1, 2.8 * cm))

    historia.append(paragrafo_pdf("JR Grupo", estilos["MarcaCapa"]))
    historia.append(paragrafo_pdf("GESTÃO DE ATIVOS DE TI", estilos["TituloCapa"]))
    historia.append(paragrafo_pdf("Relatório de Inventário de Computadores", estilos["Subtitulo"]))
    historia.append(Spacer(1, 1.1 * cm))
    capa_info = Table(
        [
            ["Empresa", "JR Serviços Empresariais"],
            ["Data de geração", data_geracao],
            ["Total de computadores", str(total)],
        ],
        colWidths=[5.2 * cm, 9.5 * cm],
    )
    capa_info.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), cor_reportlab("#E2E8F0")),
                ("TEXTCOLOR", (0, 0), (-1, -1), cor_reportlab("#0F172A")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("BOX", (0, 0), (-1, -1), 0.6, cor_reportlab("#CBD5E1")),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, cor_reportlab("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ROWBACKGROUNDS", (1, 0), (1, -1), [rl_colors.white, cor_reportlab("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    historia.append(capa_info)
    historia.append(Spacer(1, 6.2 * cm))
    historia.append(paragrafo_pdf("Sistema desenvolvido em Python + Streamlit", estilos["Subtitulo"]))
    historia.append(PageBreak())

    historia.append(paragrafo_pdf("Resumo Executivo", estilos["Secao"]))
    cards = [
        card_pdf("Total de computadores", total, cor_reportlab("#2563EB"), estilos),
        card_pdf("Ativos", total_ativos, cor_reportlab("#16A34A"), estilos),
        card_pdf("Manutenção", total_manutencao, cor_reportlab("#F97316"), estilos),
        card_pdf("Desligados", total_desligados, cor_reportlab("#DC2626"), estilos),
        card_pdf("Reserva", total_reserva, cor_reportlab("#1976D2"), estilos),
        card_pdf("Alertas", total_alertas, cor_reportlab("#EAB308"), estilos),
    ]
    historia.append(Table([cards[:3], cards[3:]], colWidths=[5.5 * cm, 5.5 * cm, 5.5 * cm], hAlign="LEFT"))
    historia.append(Spacer(1, 0.5 * cm))
    historia.append(grafico_barras_status_pdf(dados, paleta))
    historia.append(Spacer(1, 0.25 * cm))
    historia.append(grafico_barras_sala_pdf(dados))
    historia.append(Spacer(1, 0.25 * cm))
    historia.append(grafico_windows_pdf(dados))
    resumo_textual = (
        f"Foram encontrados {total} computadores distribuídos em {total_salas} setores. "
        f"A maioria encontra-se ativa. Existem {total_alertas} equipamentos que necessitam atenção."
    )
    historia.append(paragrafo_pdf(resumo_textual, estilos["Texto"]))
    historia.append(PageBreak())

    historia.append(paragrafo_pdf("Localização dos Computadores", estilos["Secao"]))
    planta_relatorio = planta_para_jpg(imagem, dados, computador_selecionado, paleta)
    planta_img = Image.open(BytesIO(planta_relatorio)).convert("RGB")
    largura_planta_pdf = doc.width
    altura_planta_pdf = min(doc.height - 3.4 * cm, largura_planta_pdf * planta_img.height / planta_img.width)
    historia.append(RLImage(BytesIO(planta_relatorio), width=largura_planta_pdf, height=altura_planta_pdf))
    historia.append(Spacer(1, 0.35 * cm))
    legenda = Table(
        [["● Ativo", "● Manutenção", "● Desligado", "● Reserva"]],
        colWidths=[4 * cm, 4 * cm, 4 * cm, 4 * cm],
    )
    legenda.setStyle(
        TableStyle(
            [
                ("TEXTCOLOR", (0, 0), (0, 0), cor_reportlab("#2E7D32")),
                ("TEXTCOLOR", (1, 0), (1, 0), cor_reportlab("#FB8C00")),
                ("TEXTCOLOR", (2, 0), (2, 0), cor_reportlab("#D32F2F")),
                ("TEXTCOLOR", (3, 0), (3, 0), cor_reportlab("#1976D2")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
            ]
        )
    )
    historia.append(legenda)
    historia.append(PageBreak())

    historia.append(paragrafo_pdf("Alertas", estilos["Secao"]))
    if alertas.empty:
        historia.append(paragrafo_pdf("Nenhum alerta encontrado no inventário atual.", estilos["Texto"]))
    else:
        for _, item in alertas.iterrows():
            motivos = ", ".join(item["motivos_alerta"])
            cor_alerta = cor_reportlab("#DC2626") if "Desligado" in motivos else cor_reportlab("#F97316")
            caixa = Table(
                [
                    [paragrafo_pdf(f"● {item['id']}", estilos["AlertaTitulo"])],
                    [paragrafo_pdf(motivos, estilos["Texto"])],
                    [paragrafo_pdf(f"Sala: {item['sala']}", estilos["Texto"])],
                ],
                colWidths=[doc.width],
            )
            caixa.setStyle(
                TableStyle(
                    [
                        ("BACKGROUND", (0, 0), (-1, -1), cor_reportlab("#FFF7ED")),
                        ("BOX", (0, 0), (-1, -1), 0.6, cor_reportlab("#FDBA74")),
                        ("LINEBEFORE", (0, 0), (0, -1), 4, cor_alerta),
                        ("LEFTPADDING", (0, 0), (-1, -1), 10),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                    ]
                )
            )
            historia.append(caixa)
            historia.append(Spacer(1, 0.22 * cm))
    historia.append(PageBreak())

    historia.append(paragrafo_pdf("Resumo dos Computadores", estilos["Secao"]))
    linhas_resumo = dados[["id", "sala", "usuario", "status", "ram", "armazenamento", "sistema"]].astype(str).values.tolist()
    historia.append(
        tabela_pdf(
            ["ID", "Sala", "Usuário", "Status", "RAM", "Armazenamento", "Sistema"],
            linhas_resumo,
            [2.0 * cm, 2.7 * cm, 2.8 * cm, 2.0 * cm, 1.5 * cm, 2.7 * cm, 2.4 * cm],
            estilos,
            fonte=7,
        )
    )
    historia.append(PageBreak())

    historia.append(paragrafo_pdf("Informações Completas", estilos["Secao"]))
    linhas_completas = dados[
        [
            "id",
            "sala",
            "usuario",
            "status",
            "sistema",
            "ram",
            "processador",
            "armazenamento",
            "placa_video",
            "observacoes",
        ]
    ].astype(str).values.tolist()
    historia.append(
        tabela_pdf(
            [
                "ID",
                "Sala",
                "Usuário",
                "Status",
                "Sistema",
                "RAM",
                "Processador",
                "Armazenamento",
                "Placa de vídeo",
                "Observações",
            ],
            linhas_completas,
            [1.3 * cm, 1.7 * cm, 1.7 * cm, 1.4 * cm, 1.8 * cm, 1.2 * cm, 2.6 * cm, 1.8 * cm, 1.8 * cm, 2.4 * cm],
            estilos,
            fonte=6.3,
        )
    )

    doc.build(historia, canvasmaker=RodapeCanvas)
    arquivo.seek(0)
    return arquivo.getvalue()


def planta_para_jpg(imagem, dados, computador_selecionado, paleta):
    jpg = imagem.copy()
    desenho = ImageDraw.Draw(jpg)

    alertas = dados[dados["status"].isin(["Manutenção", "Desligado"])]

    for _, item in alertas.iterrows():
        x = int(item["x"])
        y = int(item["y"])
        raio = 34
        desenho.ellipse(
            (x - raio, y - raio, x + raio, y + raio),
            outline=(211, 47, 47),
            width=8,
        )

    for _, item in dados.iterrows():
        x = int(item["x"])
        y = int(item["y"])
        raio = 18
        cor = cor_hex_para_rgb(paleta.get(item["status"], "#616161"))

        desenho.ellipse(
            (x - raio, y - raio, x + raio, y + raio),
            fill=cor,
            outline="white",
            width=6,
        )

    selecionado = dados[dados["id"] == computador_selecionado]

    if not selecionado.empty:
        item = selecionado.iloc[0]
        x = int(item["x"])
        y = int(item["y"])
        raio = 28
        desenho.ellipse(
            (x - raio, y - raio, x + raio, y + raio),
            outline="black",
            width=8,
        )

    arquivo = BytesIO()
    jpg.save(arquivo, format="JPEG", quality=95)
    return arquivo.getvalue()


try:
    inicializar_banco()
    importar_csv_inicial()
    df = carregar_computadores()
except Exception as erro:
    st.error("Não foi possível conectar ao PostgreSQL/Supabase.")
    st.caption(
        "Configure SUPABASE_DB_URL nas variáveis de ambiente ou em .streamlit/secrets.toml. "
        "Exemplo de chave: SUPABASE_DB_URL."
    )
    st.exception(erro)
    st.stop()

planta = Image.open(PLANTA_PATH).convert("RGB")
largura_planta, altura_planta = planta.size

df = normalizar_dataframe(df)

if df.empty:
    st.warning("Nenhum computador encontrado no banco. Importe o CSV inicial ou cadastre uma máquina.")
    st.stop()

cores = {
    "Ativo": "#2E7D32",
    "Desligado": "#D32F2F",
    "Manutenção": "#FB8C00",
    "Reserva": "#1976D2",
}

icones = {
    "Ativo": "🟢",
    "Desligado": "🔴",
    "Manutenção": "🟠",
    "Reserva": "🔵",
}

if "computador_selecionado" not in st.session_state:
    st.session_state.computador_selecionado = df.iloc[0]["id"]

if "editando_computador" not in st.session_state:
    st.session_state.editando_computador = False

if "cadastrando_computador" not in st.session_state:
    st.session_state.cadastrando_computador = False

if "status_filtrados" not in st.session_state:
    st.session_state.status_filtrados = list(cores.keys())

if "termo_busca" not in st.session_state:
    st.session_state.termo_busca = ""

if "modo_posicionar" not in st.session_state:
    st.session_state.modo_posicionar = False

if "selecionando_posicao_cadastro" not in st.session_state:
    st.session_state.selecionando_posicao_cadastro = False

if "cadastro_x" not in st.session_state:
    st.session_state.cadastro_x = min(100, largura_planta)

if "cadastro_y" not in st.session_state:
    st.session_state.cadastro_y = min(100, altura_planta)

if "mensagem_sucesso" not in st.session_state:
    st.session_state.mensagem_sucesso = ""

if "mensagem_erro" not in st.session_state:
    st.session_state.mensagem_erro = ""

if "mapa_versao" not in st.session_state:
    st.session_state.mapa_versao = 0

if st.session_state.computador_selecionado not in df["id"].tolist():
    st.session_state.computador_selecionado = df.iloc[0]["id"]

pc_parametro = st.query_params.get("pc")

if pc_parametro in df["id"].astype(str).tolist():
    st.session_state.computador_selecionado = pc_parametro

if st.session_state.mensagem_sucesso:
    st.success(st.session_state.mensagem_sucesso)
    st.session_state.mensagem_sucesso = ""

if st.session_state.mensagem_erro:
    st.error(st.session_state.mensagem_erro)
    st.session_state.mensagem_erro = ""

with st.sidebar:
    st.title("Detalhes do computador")

    if st.button("Resetar filtros", key="resetar_filtros", use_container_width=True):
        st.session_state.status_filtrados = list(cores.keys())
        st.session_state.termo_busca = ""
        st.rerun()

    status_filtrados = st.multiselect(
        "Filtrar por status",
        list(cores.keys()),
        key="status_filtrados",
    )
    termo_busca = st.text_input(
        "Buscar por nome, usuário, armazenamento ou placa de vídeo",
        placeholder="Ex.: PC-01, João, 256 GB, RTX",
        key="termo_busca",
    ).strip()
    modo_posicionar = st.checkbox(
        "Modo posicionar",
        key="modo_posicionar",
        help="Com este modo ligado, clique na planta para salvar a nova posição do computador selecionado.",
    )

    if modo_posicionar:
        st.info("Clique em um ponto da planta para reposicionar o computador selecionado.")

    df_visual = df[df["status"].isin(status_filtrados)] if status_filtrados else df.copy()

    if termo_busca:
        termo = termo_busca.lower()
        busca = (
            df_visual["id"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["usuario"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["armazenamento"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["placa_video"].astype(str).str.lower().str.contains(termo, na=False)
        )
        df_visual = df_visual[busca]

    if df_visual.empty:
        st.warning("Nenhum computador encontrado com os filtros atuais.")
        opcoes_computador = df["id"].tolist()
    else:
        opcoes_computador = df_visual["id"].tolist()

    if st.session_state.computador_selecionado not in opcoes_computador:
        st.session_state.computador_selecionado = opcoes_computador[0]

    computador_id = st.selectbox(
        "Selecione um computador",
        opcoes_computador,
        index=opcoes_computador.index(st.session_state.computador_selecionado),
    )

troca_pelo_selectbox = computador_id != st.session_state.computador_selecionado

if troca_pelo_selectbox:
    st.session_state.computador_selecionado = computador_id
    st.session_state.editando_computador = False

computador = df[df["id"] == st.session_state.computador_selecionado].iloc[0]
status_atual = computador["status"]
computadores_em_alerta = dados_com_alertas(df)
computadores_em_alerta_mapa = dados_com_alertas(df_visual)


def texto(campo):
    valor = computador.get(campo, "")
    return str(valor) if valor != "" else "Não informado"


def detalhe_linha(rotulo, valor):
    return (
        "<div class='asset-row'>"
        f"<div class='asset-label'>{escape(str(rotulo))}</div>"
        f"<div class='asset-value'>{escape(str(valor))}</div>"
        "</div>"
    )


with st.sidebar:
    st.divider()
    st.header(computador["id"])
    st.markdown(
        "<div class='asset-card'>"
        + detalhe_linha("Nome", computador["id"])
        + detalhe_linha("Status", f"{icones.get(status_atual, '⚪')} {status_atual}")
        + detalhe_linha("Sala", computador["sala"])
        + detalhe_linha("Usuário", computador["usuario"])
        + detalhe_linha("Placa de Vídeo", computador["placa_video"])
        + detalhe_linha("Armazenamento", computador["armazenamento"])
        + detalhe_linha("Sistema", computador["sistema"])
        + detalhe_linha("Memória RAM", texto("ram"))
        + detalhe_linha("Processador", texto("processador"))
        + detalhe_linha("Obs.", texto("observacoes"))
        + "</div>",
        unsafe_allow_html=True,
    )

    with st.expander("QR Code do computador"):
        st.text_input(
            "URL base do app",
            value=st.session_state.get("app_base_url", "http://localhost:8503"),
            key="app_base_url",
        )
        url_qr = url_computador(computador["id"])
        st.caption(url_qr)
        imagem_qr = qr_code_png(url_qr)

        if imagem_qr is None:
            st.warning("Instale a dependência qrcode para gerar o QR Code.")
        else:
            st.image(imagem_qr, width=180)
            st.download_button(
                "Baixar QR Code",
                data=imagem_qr,
                file_name=f"qr_{computador['id']}.png",
                mime="image/png",
                use_container_width=True,
            )

    if st.button("Editar", key="abrir_edicao", use_container_width=True):
        st.session_state.editando_computador = True
        st.session_state.cadastrando_computador = False

    if st.session_state.editando_computador:
        with st.expander("Editar máquina", expanded=True):
            with st.form(f"editar_{computador['id']}"):
                novo_id = st.text_input("Nome do computador", value=str(computador["id"]))
                novo_status = st.selectbox(
                    "Status",
                    list(cores.keys()),
                    index=list(cores.keys()).index(status_atual),
                )
                nova_sala = st.text_input("Sala", value=str(computador["sala"]))
                novo_usuario = st.text_input("Usuário", value=str(computador["usuario"]))
                nova_placa_video = st.text_input(
                    "Placa de Vídeo",
                    value=str(computador["placa_video"]),
                )
                novo_armazenamento = st.text_input("Armazenamento", value=str(computador["armazenamento"]))
                novo_sistema = st.text_input(
                    "Sistema operacional",
                    value=str(computador["sistema"]),
                )
                nova_ram = st.text_input("Memória RAM", value=str(computador["ram"]))
                novo_processador = st.text_input(
                    "Processador",
                    value=str(computador["processador"]),
                )
                novas_observacoes = st.text_area(
                    "Observações",
                    value=str(computador["observacoes"]),
                )
                novo_x = st.number_input(
                    "Coordenada X",
                    min_value=0,
                    max_value=largura_planta,
                    value=int(computador["x"]),
                    step=1,
                )
                novo_y = st.number_input(
                    "Coordenada Y",
                    min_value=0,
                    max_value=altura_planta,
                    value=int(computador["y"]),
                    step=1,
                )

                salvar = st.form_submit_button("Salvar", use_container_width=True)
                cancelar = st.form_submit_button("Cancelar", use_container_width=True)

            if cancelar:
                st.session_state.editando_computador = False
                st.rerun()

            if salvar:
                registro_atualizado = {
                    "id": novo_id.strip(),
                    "sala": nova_sala.strip(),
                    "x": int(novo_x),
                    "y": int(novo_y),
                    "status": novo_status,
                    "armazenamento": novo_armazenamento.strip(),
                    "placa_video": nova_placa_video.strip(),
                    "usuario": novo_usuario.strip(),
                    "sistema": novo_sistema.strip(),
                    "ram": nova_ram.strip(),
                    "processador": novo_processador.strip(),
                    "observacoes": novas_observacoes.strip(),
                }
                erros = validar_registro(
                    registro_atualizado,
                    df,
                    id_original=computador["id"],
                )

                if erros:
                    for erro in erros:
                        st.error(erro)
                else:
                    try:
                        atualizou = atualizar_computador(
                            computador["id"],
                            registro_atualizado,
                            "Edição",
                        )
                    except Exception as erro:
                        st.error(f"Não foi possível atualizar no Supabase: {erro}")
                    else:
                        if atualizou:
                            st.session_state.computador_selecionado = registro_atualizado["id"]
                            st.session_state.editando_computador = False
                            st.success("Dados atualizados.")
                            st.rerun()
                        else:
                            st.error("Computador não encontrado no banco.")


with st.sidebar:
    if not computadores_em_alerta.empty:
        st.divider()
        with st.expander("Indicador de alerta", expanded=True):
            for _, item in computadores_em_alerta.iterrows():
                motivos = ", ".join(item["motivos_alerta"])
                st.warning(
                    f"{icones.get(item['status'], '⚠️')} {item['id']} - "
                    f"{item['sala']} | {motivos}"
                )

    st.divider()
    with st.expander("Exportar inventário"):
        st.download_button(
            "Baixar CSV",
            data=df.to_csv(index=False).encode("utf-8-sig"),
            file_name="inventario_computadores.csv",
            mime="text/csv",
            use_container_width=True,
        )
        st.download_button(
            "Baixar Excel",
            data=dataframe_para_xlsx(df),
            file_name="inventario_computadores.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Baixar PDF",
            data=relatorio_pdf(
                planta,
                df,
                computadores_em_alerta,
                cores,
                st.session_state.computador_selecionado,
            ),
            file_name="relatorio_inventario.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.download_button(
            "Baixar planta JPG",
            data=planta_para_jpg(
                planta,
                df,
                st.session_state.computador_selecionado,
                cores,
            ),
            file_name="planta_computadores.jpg",
            mime="image/jpeg",
            use_container_width=True,
        )

    with st.expander("Cadastro de computador", expanded=st.session_state.cadastrando_computador):

        if st.button(
            "Cadastrar computador",
            key="abrir_cadastro",
            use_container_width=True,
        ):
            st.session_state.cadastrando_computador = True
            st.session_state.editando_computador = False
            st.session_state.selecionando_posicao_cadastro = False
            st.session_state.cadastro_x = min(100, largura_planta)
            st.session_state.cadastro_y = min(100, altura_planta)

        if st.session_state.cadastrando_computador:
            proximo_id = f"PC-{len(df) + 1:02d}"

            if st.button(
                "Selecionar posição na planta",
                key="selecionar_posicao_cadastro",
                use_container_width=True,
            ):
                st.session_state.selecionando_posicao_cadastro = not st.session_state.selecionando_posicao_cadastro

            if st.session_state.selecionando_posicao_cadastro:
                st.info("Clique na planta para preencher as coordenadas do novo computador.")
                coordenadas_cadastro = streamlit_image_coordinates(
                    planta,
                    key=f"planta_cadastro_{st.session_state.mapa_versao}",
                )

                if coordenadas_cadastro:
                    st.session_state.cadastro_x = max(
                        0,
                        min(int(round(coordenadas_cadastro["x"])), largura_planta),
                    )
                    st.session_state.cadastro_y = max(
                        0,
                        min(int(round(coordenadas_cadastro["y"])), altura_planta),
                    )
                    st.session_state.selecionando_posicao_cadastro = False
                    st.success(
                        f"Posição selecionada: X {st.session_state.cadastro_x}, Y {st.session_state.cadastro_y}."
                    )

            with st.form("cadastrar_computador"):
                cadastro_id = st.text_input("Nome do computador", value=proximo_id)
                cadastro_status = st.selectbox("Status", list(cores.keys()))
                cadastro_sala = st.text_input("Sala")
                cadastro_usuario = st.text_input("Usuário")
                cadastro_placa_video = st.text_input("Placa de Vídeo")
                cadastro_armazenamento = st.text_input("Armazenamento")
                cadastro_sistema = st.text_input(
                    "Sistema operacional",
                    value="Windows 11",
                )
                cadastro_ram = st.text_input("Memória RAM", value="16 GB")
                cadastro_processador = st.text_input("Processador")
                cadastro_observacoes = st.text_area("Observações")
                cadastro_x = st.number_input(
                    "Coordenada X",
                    min_value=0,
                    max_value=largura_planta,
                    value=int(st.session_state.cadastro_x),
                    step=1,
                    key="cadastro_x_input",
                )
                cadastro_y = st.number_input(
                    "Coordenada Y",
                    min_value=0,
                    max_value=altura_planta,
                    value=int(st.session_state.cadastro_y),
                    step=1,
                    key="cadastro_y_input",
                )

                cadastrar = st.form_submit_button(
                    "Salvar cadastro",
                    use_container_width=True,
                )
                cancelar_cadastro = st.form_submit_button(
                    "Cancelar",
                    use_container_width=True,
                )

            if cancelar_cadastro:
                st.session_state.cadastrando_computador = False
                st.session_state.selecionando_posicao_cadastro = False
                st.session_state.cadastro_x = min(100, largura_planta)
                st.session_state.cadastro_y = min(100, altura_planta)
                st.rerun()

            if cadastrar:
                novo_computador = {
                    "id": cadastro_id.strip(),
                    "sala": cadastro_sala.strip(),
                    "x": int(cadastro_x),
                    "y": int(cadastro_y),
                    "status": cadastro_status,
                    "armazenamento": cadastro_armazenamento.strip(),
                    "placa_video": cadastro_placa_video.strip(),
                    "usuario": cadastro_usuario.strip(),
                    "sistema": cadastro_sistema.strip(),
                    "ram": cadastro_ram.strip(),
                    "processador": cadastro_processador.strip(),
                    "observacoes": cadastro_observacoes.strip(),
                }
                erros = validar_registro(novo_computador, df)

                if erros:
                    for erro in erros:
                        st.error(erro)
                else:
                    try:
                        salvar_computador(novo_computador)
                    except Exception as erro:
                        st.error(f"Não foi possível cadastrar no Supabase: {erro}")
                    else:
                        st.session_state.computador_selecionado = novo_computador["id"]
                        st.session_state.cadastrando_computador = False
                        st.session_state.selecionando_posicao_cadastro = False
                        st.session_state.cadastro_x = min(100, largura_planta)
                        st.session_state.cadastro_y = min(100, altura_planta)
                        st.success("Computador cadastrado.")
                        st.rerun()

    with st.expander("Importação em massa"):
        arquivo_importacao = st.file_uploader(
            "Importar CSV",
            type=["csv"],
            key="arquivo_importacao_csv",
        )

        if arquivo_importacao is not None:
            if st.button(
                "Importar computadores",
                key="importar_computadores",
                use_container_width=True,
            ):
                try:
                    dados_importados = pd.read_csv(arquivo_importacao).fillna("")
                    dados_importados = normalizar_dataframe(dados_importados)
                except Exception as erro:
                    st.error(f"Não foi possível ler o CSV: {erro}")
                    dados_importados = pd.DataFrame()

                if not dados_importados.empty:
                    erros_importacao = []
                    df_atualizado = df.copy()
                    registros_validos = []

                    for numero_linha, item in dados_importados.iterrows():
                        registro = item.to_dict()
                        erros = validar_registro(registro, df_atualizado)

                        if erros:
                            erros_importacao.append(
                                f"Linha {numero_linha + 2} ({registro.get('id', '')}): "
                                + " ".join(erros)
                            )
                            continue

                        df_atualizado = pd.concat(
                            [df_atualizado, pd.DataFrame([registro])],
                            ignore_index=True,
                        )
                        registros_validos.append(registro)

                    if erros_importacao:
                        for erro in erros_importacao:
                            st.error(erro)

                    if registros_validos:
                        try:
                            importar_computadores(registros_validos)
                        except Exception as erro:
                            st.error(f"Não foi possível importar no Supabase: {erro}")
                        else:
                            st.success(
                                f"{len(registros_validos)} computador(es) importado(s)."
                            )
                            st.rerun()

    with st.expander("Controle de movimentação"):
        try:
            historico_movimentacao = carregar_movimentacoes()
        except Exception as erro:
            st.error(f"Não foi possível carregar o histórico do Supabase: {erro}")
            historico_movimentacao = pd.DataFrame(columns=CAMPOS_MOVIMENTACAO)

        if not historico_movimentacao.empty:
            st.dataframe(
                historico_movimentacao.tail(10),
                hide_index=True,
                use_container_width=True,
            )
            st.download_button(
                "Baixar histórico",
                data=historico_movimentacao.to_csv(index=False).encode("utf-8-sig"),
                file_name="historico_movimentacoes.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.info("Nenhuma movimentação registrada ainda.")

    with st.expander("Backups automáticos"):
        backups = sorted(BACKUP_DIR.glob("computadores_*.csv"), reverse=True)

        if backups:
            for backup in backups[:5]:
                st.caption(backup.name)
        else:
            st.info("Nenhum backup criado ainda.")

st.markdown(
    f"""
    <div style="
        text-align: center;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        margin: 0.4rem 0 1.8rem 0;
    ">
        {logo_html_dashboard()}
        <div style="font-size: 30px; font-weight: 700; color: #F8FAFC; line-height: 1.15;">JR Grupo</div>
        <div style="font-size: 20px; font-weight: 600; color: #CBD5E1; margin-top: 6px; line-height: 1.25;">Gestão de Ativos de TI</div>
        <div style="font-size: 16px; color: #94A3B8; margin-top: 4px; line-height: 1.3;">Dashboard de Inventário de Computadores</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown("### Dashboard")
total_computadores = len(df)
total_ativos = int((df["status"] == "Ativo").sum())
total_manutencao = int((df["status"] == "Manutenção").sum())
total_desligados = int((df["status"] == "Desligado").sum())
total_reservas = int((df["status"] == "Reserva").sum())
total_alertas = len(computadores_em_alerta)
inventario_completo = percentual_inventario_completo(df)
pendencias_df = tabela_pendencias(df)

cards_dashboard = [
    ("Total de computadores", total_computadores, "#3B82F6"),
    ("Ativos", total_ativos, "#22C55E"),
    ("Em manutenção", total_manutencao, "#F97316"),
    ("Desligados", total_desligados, "#EF4444"),
    ("Reservas", total_reservas, "#60A5FA"),
    ("Com alerta", total_alertas, "#EAB308"),
    ("Inventário completo", f"{inventario_completo}%", "#14B8A6"),
]
st.markdown(
    "<div class='dash-card-grid'>"
    + "".join(render_dashboard_card(titulo, valor, cor) for titulo, valor, cor in cards_dashboard)
    + "</div>",
    unsafe_allow_html=True,
)

status_mais_comum = df["status"].value_counts().idxmax() if not df.empty else "sem status"
frase_status = (
    "A maior parte está ativa."
    if status_mais_comum == "Ativo"
    else f"A maior parte está com status {status_mais_comum}."
)
resumo_executivo = (
    f"O inventário possui {total_computadores} computadores cadastrados. "
    f"{frase_status} Existem {total_alertas} equipamentos com alerta e "
    f"o inventário está {inventario_completo}% completo."
)
st.markdown(f"<div class='executive-summary'>{escape(resumo_executivo)}</div>", unsafe_allow_html=True)

grafico_layout = dict(
    template="plotly_dark",
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(15,23,42,0.35)",
    font=dict(color="#E2E8F0"),
    height=330,
    margin=dict(l=10, r=10, t=52, b=20),
)

dash_col1, dash_col2 = st.columns(2)

status_dashboard = df["status"].value_counts().reindex(list(cores.keys()), fill_value=0).reset_index()
status_dashboard.columns = ["status", "total"]
fig_status = go.Figure(
    go.Bar(
        x=status_dashboard["status"],
        y=status_dashboard["total"],
        marker_color=[cores.get(status, "#616161") for status in status_dashboard["status"]],
        text=status_dashboard["total"],
        textposition="outside",
    )
)
fig_status.update_layout(title="Total por status", **grafico_layout)
fig_status.update_yaxes(gridcolor="rgba(148,163,184,0.16)")
dash_col1.plotly_chart(fig_status, use_container_width=True)

ordem_ram = ["4 GB", "6 GB", "8 GB", "16 GB", "32 GB", "A coletar", "Outros"]
ram_dashboard = (
    df["ram"]
    .map(padronizar_ram)
    .value_counts()
    .reindex(ordem_ram, fill_value=0)
    .reset_index()
)
ram_dashboard.columns = ["ram", "total"]
ram_dashboard = ram_dashboard[(ram_dashboard["total"] > 0) | (ram_dashboard["ram"] != "Outros")]
fig_ram = go.Figure(
    go.Bar(
        x=ram_dashboard["ram"],
        y=ram_dashboard["total"],
        marker_color="#14B8A6",
        text=ram_dashboard["total"],
        textposition="outside",
    )
)
fig_ram.update_layout(title="Quantidade por memória RAM", **grafico_layout)
fig_ram.update_yaxes(gridcolor="rgba(148,163,184,0.16)")
dash_col2.plotly_chart(fig_ram, use_container_width=True)

dash_col3, dash_col4 = st.columns(2)

sistema_dashboard = (
    df["sistema"]
    .map(padronizar_sistema)
    .value_counts()
    .reindex(["Windows 10", "Windows 11", "Outros"], fill_value=0)
    .reset_index()
)
sistema_dashboard.columns = ["sistema", "total"]
fig_sistema = go.Figure(
    go.Pie(
        labels=sistema_dashboard["sistema"],
        values=sistema_dashboard["total"],
        hole=0.48,
        marker=dict(colors=["#60A5FA", "#22C55E", "#A78BFA"]),
    )
)
fig_sistema.update_layout(title="Sistema operacional", **grafico_layout)
dash_col3.plotly_chart(fig_sistema, use_container_width=True)

alertas_dashboard = contar_alertas_por_tipo(df)
fig_alertas = go.Figure(
    go.Bar(
        x=alertas_dashboard["total"],
        y=alertas_dashboard["tipo"],
        orientation="h",
        marker_color="#F97316",
        text=alertas_dashboard["total"],
        textposition="auto",
    )
)
fig_alertas.update_layout(title="Alertas por tipo", **grafico_layout)
fig_alertas.update_yaxes(autorange="reversed", gridcolor="rgba(148,163,184,0.08)")
fig_alertas.update_xaxes(gridcolor="rgba(148,163,184,0.16)")
dash_col4.plotly_chart(fig_alertas, use_container_width=True)

dash_col5, _ = st.columns(2)
sala_dashboard = df["sala"].replace("", "Não informado").value_counts().head(10).reset_index()
sala_dashboard.columns = ["sala", "total"]
fig_sala = go.Figure(
    go.Bar(
        x=sala_dashboard["total"],
        y=sala_dashboard["sala"],
        orientation="h",
        marker_color="#3B82F6",
        text=sala_dashboard["total"],
        textposition="auto",
    )
)
fig_sala.update_layout(title="Top salas por quantidade", **grafico_layout)
fig_sala.update_yaxes(autorange="reversed", gridcolor="rgba(148,163,184,0.08)")
fig_sala.update_xaxes(gridcolor="rgba(148,163,184,0.16)")
dash_col5.plotly_chart(fig_sala, use_container_width=True)

st.markdown("### Pendências do inventário")
if pendencias_df.empty:
    st.success("Nenhuma pendência encontrada nos principais campos do inventário.")
else:
    st.dataframe(pendencias_df, hide_index=True, use_container_width=True)

st.divider()
st.markdown("### Planta de localização")
st.caption("Use a planta para localizar e selecionar computadores. Os detalhes completos ficam na sidebar.")

total = len(df_visual)
ativos = (df_visual["status"] == "Ativo").sum()
manutencao = (df_visual["status"] == "Manutenção").sum()
desligados = (df_visual["status"] == "Desligado").sum()
reservas = (df_visual["status"] == "Reserva").sum()

fig = go.Figure()
ids_por_trace = {}

fig.add_layout_image(
    dict(
        source=planta,
        xref="x",
        yref="y",
        x=0,
        y=0,
        sizex=largura_planta,
        sizey=altura_planta,
        sizing="stretch",
        xanchor="left",
        yanchor="top",
        layer="below",
    )
)

if not computadores_em_alerta_mapa.empty:
    ids_por_trace[len(fig.data)] = computadores_em_alerta_mapa["id"].tolist()

    fig.add_trace(
        go.Scatter(
            x=computadores_em_alerta_mapa["x"].tolist(),
            y=computadores_em_alerta_mapa["y"].tolist(),
            mode="markers",
            name="Alerta",
            showlegend=False,
            marker=dict(
                size=16,
                color="rgba(0,0,0,0)",
                line=dict(color="rgba(211,47,47,0.55)", width=2),
            ),
            hoverinfo="skip",
            hovertemplate=None,
        )
    )

for status in ["Ativo", "Desligado", "Manutenção", "Reserva"]:
    grupo = df_visual[df_visual["status"] == status]

    if grupo.empty:
        continue

    ids_por_trace[len(fig.data)] = grupo["id"].tolist()

    fig.add_trace(
        go.Scatter(
            x=grupo["x"].tolist(),
            y=grupo["y"].tolist(),
            mode="markers",
            name=status,
            showlegend=False,
            marker=dict(
                size=8,
                color=cores.get(status, "#616161"),
                line=dict(color="rgba(255,255,255,0.9)", width=1),
            ),
            customdata=list(
                zip(
                    grupo["id"],
                    grupo["sala"],
                    grupo["status"].map(lambda valor: f"{icones.get(valor, '')} {valor}"),
                )
            ),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "%{customdata[1]}<br>"
                "%{customdata[2]}"
                "<extra></extra>"
            ),
        )
    )

ids_por_trace[len(fig.data)] = [computador["id"]]

fig.add_trace(
    go.Scatter(
        x=[computador["x"]],
        y=[computador["y"]],
        mode="markers",
        showlegend=False,
        marker=dict(
            size=18,
            color="rgba(255,255,255,0)",
            line=dict(color="rgba(25,118,210,0.35)", width=3),
        ),
        hoverinfo="skip",
        hovertemplate=None,
    )
)

fig.add_trace(
    go.Scatter(
        x=[computador["x"]],
        y=[computador["y"]],
        mode="markers",
        showlegend=False,
        marker=dict(
            size=12,
            color=cores.get(computador["status"], "#616161"),
            line=dict(color="white", width=3),
        ),
        hoverinfo="skip",
        hovertemplate=None,
    )
)

fig.add_annotation(
    text="🟢 Ativo&nbsp;&nbsp;🟠 Manutenção&nbsp;&nbsp;🔴 Desligado&nbsp;&nbsp;🔵 Reserva",
    xref="paper",
    yref="paper",
    x=0.02,
    y=1.045,
    xanchor="left",
    yanchor="top",
    showarrow=False,
    align="left",
    font=dict(size=12, color="#111827"),
)

fig.update_layout(
    autosize=True,
    height=820,
    coloraxis_showscale=False,
    showlegend=False,
    hovermode="closest",
    clickmode="event+select",
    hoverdistance=5,
    spikedistance=-1,
    xaxis=dict(
        visible=False,
        range=[0, largura_planta - 2],
        showgrid=False,
        zeroline=False,
        showline=False,
        mirror=False,
        fixedrange=True,
        constrain="domain",
    ),
    yaxis=dict(
        visible=False,
        range=[altura_planta, 0],
        showgrid=False,
        zeroline=False,
        showline=False,
        mirror=False,
        fixedrange=True,
        constrain="domain",
        scaleanchor="x",
        scaleratio=1,
    ),
    hoverlabel=dict(
        bgcolor="white",
        bordercolor="rgba(25,118,210,0.45)",
        font_size=12,
        font_family="Arial",
        font_color="black",
        align="left",
    ),
    margin=dict(l=0, r=0, t=0, b=0),
)

if modo_posicionar:
    coordenadas_imagem = streamlit_image_coordinates(
        planta,
        key=f"planta_posicionar_{st.session_state.mapa_versao}",
    )

    if coordenadas_imagem:
        computador_id = st.session_state.computador_selecionado
        novo_x = int(round(coordenadas_imagem["x"]))
        novo_y = int(round(coordenadas_imagem["y"]))
        novo_x = max(0, min(novo_x, largura_planta))
        novo_y = max(0, min(novo_y, altura_planta))

        try:
            df, salvou_posicao = salvar_posicao_computador(df, computador_id, novo_x, novo_y)
        except Exception as erro:
            salvou_posicao = False
            st.session_state.mensagem_erro = f"Não foi possível salvar no Supabase: {erro}"

        st.session_state.computador_selecionado = computador_id

        if salvou_posicao:
            st.session_state.mensagem_sucesso = f"{computador_id} reposicionado."
        elif not st.session_state.mensagem_erro:
            st.session_state.mensagem_erro = "Não foi possível salvar a nova posição."

        st.session_state.mapa_versao += 1
        st.rerun()
else:
    selected_points = plotly_events(
        fig,
        click_event=True,
        hover_event=False,
        select_event=False,
        override_height=820,
        key=f"mapa_ativos_{st.session_state.mapa_versao}",
    )

    if selected_points and not troca_pelo_selectbox:
        novo_computador_id = computador_clicado(selected_points, ids_por_trace, df)

        if (
            novo_computador_id
            and st.session_state.computador_selecionado != novo_computador_id
        ):
            st.session_state.computador_selecionado = novo_computador_id
            st.session_state.editando_computador = False
            st.rerun()
