from pathlib import Path
from datetime import datetime
from io import BytesIO
from shutil import copy2
from textwrap import shorten
from unicodedata import normalize
from zipfile import ZIP_DEFLATED, ZipFile
from xml.sax.saxutils import escape

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from PIL import Image, ImageDraw
from streamlit_plotly_events import plotly_events

BASE_DIR = Path(__file__).resolve().parent

PLANTA_PATH = BASE_DIR / "assets" / "planta.png"
# Planta real reconstruída pelo contorno do prédio.
# Usar quando for trocar a base do sistema:
# PLANTA_PATH = BASE_DIR / "assets" / "planta_real.png"

CSV_PATH = BASE_DIR / "data" / "computadores.csv"
BACKUP_DIR = BASE_DIR / "data" / "backups"
MOVIMENTACOES_PATH = BASE_DIR / "data" / "movimentacoes.csv"

CAMPOS_CSV = [
    "id",
    "sala",
    "x",
    "y",
    "status",
    "ip",
    "patrimonio",
    "usuario",
    "sistema",
    "ram",
    "processador",
    "observacoes",
]
CAMPOS_MOVIMENTACAO = [
    "data_hora",
    "id",
    "campo",
    "valor_anterior",
    "valor_novo",
    "acao",
]
CAMPOS_MOVIMENTACAO_MONITORADOS = ["sala", "usuario", "x", "y"]
SISTEMAS_ANTIGOS = ["Windows 7", "Windows 8", "Windows 8.1", "Windows 10"]
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


def agora_id():
    return datetime.now().strftime("%Y%m%d_%H%M%S_%f")


def agora_texto():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def criar_backup_csv():
    if not CSV_PATH.exists():
        return None

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    destino = BACKUP_DIR / f"computadores_{agora_id()}.csv"
    copy2(CSV_PATH, destino)
    return destino


def salvar_com_backup(dados):
    criar_backup_csv()
    dados = dados.reindex(columns=CAMPOS_CSV)
    dados.to_csv(CSV_PATH, index=False)


def normalizar_dataframe(dados):
    dados = dados.copy()

    for coluna in CAMPOS_CSV:
        if coluna not in dados.columns:
            dados[coluna] = ""

    dados = dados.reindex(columns=CAMPOS_CSV).fillna("")

    for coluna in ["x", "y"]:
        dados[coluna] = pd.to_numeric(dados[coluna], errors="coerce").fillna(0).astype(int)

    return dados


def valor_limpo(valor):
    return str(valor).strip()


def texto_pdf(valor):
    texto = normalize("NFKD", str(valor)).encode("ascii", "ignore").decode("ascii")
    return texto


def validar_registro(registro, dados_base, id_original=None):
    erros = []
    computador_id = valor_limpo(registro.get("id", ""))
    ip = valor_limpo(registro.get("ip", ""))
    patrimonio = valor_limpo(registro.get("patrimonio", ""))
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

    if ip and ip in comparacao["ip"].astype(str).tolist():
        erros.append("Já existe um computador com esse IP.")

    if patrimonio and patrimonio in comparacao["patrimonio"].astype(str).tolist():
        erros.append("Já existe um computador com esse patrimônio.")

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
                    "id": registro_novo.get("id", registro_antigo.get("id", "")),
                    "campo": campo,
                    "valor_anterior": valor_anterior,
                    "valor_novo": valor_novo,
                    "acao": acao,
                }
            )

    if not linhas:
        return

    if MOVIMENTACOES_PATH.exists():
        historico = pd.read_csv(MOVIMENTACOES_PATH).fillna("")
    else:
        historico = pd.DataFrame(columns=CAMPOS_MOVIMENTACAO)

    historico = pd.concat([historico, pd.DataFrame(linhas)], ignore_index=True)
    historico = historico.reindex(columns=CAMPOS_MOVIMENTACAO)
    historico.to_csv(MOVIMENTACOES_PATH, index=False)


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
        motivos.append("sem usuário")

    if ram.startswith("8 gb") or ram.startswith("4 gb"):
        motivos.append("pouca RAM")

    if sistema in SISTEMAS_ANTIGOS:
        motivos.append("sistema antigo")

    if any(palavra in observacoes.lower() for palavra in PALAVRAS_CRITICAS):
        motivos.append("observação crítica")

    return motivos


def dados_com_alertas(dados):
    dados = dados.copy()
    dados["motivos_alerta"] = dados.apply(motivos_alerta, axis=1)
    return dados[dados["motivos_alerta"].map(bool)]


def relatorio_pdf(imagem, dados, alertas, paleta, computador_selecionado):
    planta_relatorio = planta_para_jpg(imagem, dados, computador_selecionado, paleta)
    planta_img = Image.open(BytesIO(planta_relatorio)).convert("RGB")
    largura = 1600
    planta_altura = int(planta_img.height * 760 / planta_img.width)
    planta_img = planta_img.resize((760, planta_altura))
    altura = max(1200, 420 + planta_altura + (len(dados) * 28))
    pagina = Image.new("RGB", (largura, altura), "white")
    desenho = ImageDraw.Draw(pagina)

    y = 40
    desenho.text((50, y), texto_pdf("Relatório do inventário de computadores"), fill=(17, 24, 39))
    y += 36
    desenho.text(
        (50, y),
        texto_pdf(
            f"Total: {len(dados)} | Ativos: {(dados['status'] == 'Ativo').sum()} | "
            f"Manutenção: {(dados['status'] == 'Manutenção').sum()} | "
            f"Desligados: {(dados['status'] == 'Desligado').sum()} | "
            f"Reservas: {(dados['status'] == 'Reserva').sum()}"
        ),
        fill=(17, 24, 39),
    )
    y += 44
    desenho.text((50, y), "Alertas", fill=(17, 24, 39))
    y += 26

    if alertas.empty:
        desenho.text((70, y), "Nenhum alerta encontrado.", fill=(34, 197, 94))
        y += 28
    else:
        for _, item in alertas.iterrows():
            motivos = ", ".join(item["motivos_alerta"])
            texto_alerta = f"{item['id']} - {item['sala']} - {motivos}"
            desenho.text((70, y), texto_pdf(shorten(texto_alerta, width=115)), fill=(180, 35, 24))
            y += 26

    y += 20
    pagina.paste(planta_img, (50, y))
    y += planta_img.height + 36
    desenho.text((50, y), "Inventario", fill=(17, 24, 39))
    y += 30

    for _, item in dados.iterrows():
        linha = (
            f"{item['id']} | {item['status']} | {item['sala']} | "
            f"{item['usuario']} | {item['ip']} | {item['patrimonio']}"
        )
        desenho.text((70, y), texto_pdf(shorten(linha, width=150)), fill=(17, 24, 39))
        y += 26

    arquivo = BytesIO()
    pagina.save(arquivo, format="PDF", resolution=100)
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


df = pd.read_csv(CSV_PATH)
planta = Image.open(PLANTA_PATH).convert("RGB")
largura_planta, altura_planta = planta.size

df = normalizar_dataframe(df)

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

if st.session_state.computador_selecionado not in df["id"].tolist():
    st.session_state.computador_selecionado = df.iloc[0]["id"]

with st.sidebar:
    st.title("Detalhes do computador")

    status_filtrados = st.multiselect(
        "Filtrar por status",
        list(cores.keys()),
        default=list(cores.keys()),
    )
    termo_busca = st.text_input(
        "Buscar por usuário, IP ou patrimônio",
        placeholder="Ex.: João, 192.168.0.11, TI001",
    ).strip()

    df_visual = df[df["status"].isin(status_filtrados)] if status_filtrados else df.copy()

    if termo_busca:
        termo = termo_busca.lower()
        busca = (
            df_visual["id"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["usuario"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["ip"].astype(str).str.lower().str.contains(termo, na=False)
            | df_visual["patrimonio"].astype(str).str.lower().str.contains(termo, na=False)
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
        + detalhe_linha("Patrimônio", computador["patrimonio"])
        + detalhe_linha("IP", computador["ip"])
        + detalhe_linha("Sistema", computador["sistema"])
        + detalhe_linha("RAM", texto("ram"))
        + detalhe_linha("CPU", texto("processador"))
        + detalhe_linha("Obs.", texto("observacoes"))
        + "</div>",
        unsafe_allow_html=True,
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
                novo_patrimonio = st.text_input(
                    "Patrimônio",
                    value=str(computador["patrimonio"]),
                )
                novo_ip = st.text_input("IP", value=str(computador["ip"]))
                novo_sistema = st.text_input(
                    "Sistema operacional",
                    value=str(computador["sistema"]),
                )
                nova_ram = st.text_input("RAM", value=str(computador["ram"]))
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
                    "ip": novo_ip.strip(),
                    "patrimonio": novo_patrimonio.strip(),
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
                    indice = df.index[df["id"] == computador["id"]][0]
                    registro_antigo = df.loc[indice].to_dict()

                    for campo, valor in registro_atualizado.items():
                        df.loc[indice, campo] = valor

                    registrar_movimentacoes(
                        registro_antigo,
                        registro_atualizado,
                        "Edição",
                    )
                    salvar_com_backup(df)

                    st.session_state.computador_selecionado = registro_atualizado["id"]
                    st.session_state.editando_computador = False
                    st.success("Dados atualizados.")
                    st.rerun()


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

        if st.session_state.cadastrando_computador:
            proximo_id = f"PC-{len(df) + 1:02d}"

            with st.form("cadastrar_computador"):
                cadastro_id = st.text_input("Nome do computador", value=proximo_id)
                cadastro_status = st.selectbox("Status", list(cores.keys()))
                cadastro_sala = st.text_input("Sala")
                cadastro_usuario = st.text_input("Usuário")
                cadastro_patrimonio = st.text_input("Patrimônio")
                cadastro_ip = st.text_input("IP")
                cadastro_sistema = st.text_input(
                    "Sistema operacional",
                    value="Windows 11",
                )
                cadastro_ram = st.text_input("RAM", value="16 GB")
                cadastro_processador = st.text_input("Processador")
                cadastro_observacoes = st.text_area("Observações")
                cadastro_x = st.number_input(
                    "Coordenada X",
                    min_value=0,
                    max_value=largura_planta,
                    value=min(100, largura_planta),
                    step=1,
                )
                cadastro_y = st.number_input(
                    "Coordenada Y",
                    min_value=0,
                    max_value=altura_planta,
                    value=min(100, altura_planta),
                    step=1,
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
                st.rerun()

            if cadastrar:
                novo_computador = {
                    "id": cadastro_id.strip(),
                    "sala": cadastro_sala.strip(),
                    "x": int(cadastro_x),
                    "y": int(cadastro_y),
                    "status": cadastro_status,
                    "ip": cadastro_ip.strip(),
                    "patrimonio": cadastro_patrimonio.strip(),
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
                    df_atualizado = pd.concat(
                        [df, pd.DataFrame([novo_computador])],
                        ignore_index=True,
                    )
                    salvar_com_backup(df_atualizado)

                    st.session_state.computador_selecionado = novo_computador["id"]
                    st.session_state.cadastrando_computador = False
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

                    if erros_importacao:
                        for erro in erros_importacao:
                            st.error(erro)

                    if len(df_atualizado) > len(df):
                        salvar_com_backup(df_atualizado)
                        st.success(
                            f"{len(df_atualizado) - len(df)} computador(es) importado(s)."
                        )
                        st.rerun()

    with st.expander("Controle de movimentação"):
        if MOVIMENTACOES_PATH.exists():
            historico_movimentacao = pd.read_csv(MOVIMENTACOES_PATH).fillna("")
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
                size=36,
                color="rgba(0,0,0,0)",
                line=dict(color="#D32F2F", width=4),
            ),
            hoverinfo="skip",
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
                size=18,
                color=cores.get(status, "#616161"),
                line=dict(color="white", width=2),
            ),
            customdata=grupo[
                ["id", "sala", "ip", "patrimonio", "status", "usuario", "sistema"]
            ].values.tolist(),
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "Sala: %{customdata[1]}<br>"
                "IP: %{customdata[2]}<br>"
                "Patrimônio: %{customdata[3]}<br>"
                "Status: %{customdata[4]}"
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
            size=42,
            color="rgba(0,0,0,0)",
            line=dict(color="black", width=5),
        ),
        hoverinfo="skip",
    )
)

fig.add_annotation(
    text="🟢 Ativo&nbsp;&nbsp;&nbsp;🔴 Desligado&nbsp;&nbsp;&nbsp;🟠 Manutenção&nbsp;&nbsp;&nbsp;🔵 Reserva",
    xref="paper",
    yref="paper",
    x=0.02,
    y=1.06,
    xanchor="left",
    yanchor="top",
    showarrow=False,
    align="left",
    font=dict(size=13, color="#111827"),
)

fig.update_layout(
    title=dict(
        text=(
            "Inventário de Computadores - Planta Interativa"
            f"<br><sup>Total: {total} | Ativos: {ativos} | "
            f"Manutenção: {manutencao} | Desligados: {desligados} | Reservas: {reservas}</sup>"
        ),
        x=0.02,
        xanchor="left",
    ),
    height=820,
    coloraxis_showscale=False,
    showlegend=False,
    paper_bgcolor="white",
    plot_bgcolor="white",
    hoverlabel=dict(
        bgcolor="white",
        bordercolor="#1976D2",
        font_size=13,
        font_family="Arial",
        font_color="black",
    ),
    margin=dict(l=10, r=10, t=110, b=10),
)

fig.update_xaxes(
    visible=False,
    range=[0, largura_planta],
    showgrid=False,
    zeroline=False,
)
fig.update_yaxes(
    visible=False,
    range=[altura_planta, 0],
    showgrid=False,
    zeroline=False,
    scaleanchor="x",
    scaleratio=1,
)

selected_points = plotly_events(
    fig,
    click_event=True,
    hover_event=False,
    select_event=False,
    override_height=820,
    override_width=1400,
    key=f"mapa_ativos_{st.session_state.computador_selecionado}",
)

if selected_points and not troca_pelo_selectbox:
    novo_computador_id = None

    for ponto in selected_points:
        numero_trace = ponto.get("curveNumber")
        numero_ponto = ponto.get("pointNumber")

        if numero_trace is not None:
            numero_trace = int(numero_trace)

        if numero_ponto is not None:
            numero_ponto = int(numero_ponto)

        ids_do_trace = ids_por_trace.get(numero_trace, [])

        if numero_ponto is not None and numero_ponto < len(ids_do_trace):
            novo_computador_id = ids_do_trace[numero_ponto]
            break

    if novo_computador_id is None:
        ponto = selected_points[0]
        ponto_x = ponto.get("x")
        ponto_y = ponto.get("y")
        computador_clicado = df[(df["x"] == ponto_x) & (df["y"] == ponto_y)]

        if not computador_clicado.empty:
            novo_computador_id = computador_clicado.iloc[0]["id"]

    if (
        novo_computador_id
        and st.session_state.computador_selecionado != novo_computador_id
    ):
        st.session_state.computador_selecionado = novo_computador_id
        st.rerun()
