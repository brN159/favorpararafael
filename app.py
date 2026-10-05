import io
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

import tse

PASTA = Path("dados")
UFS = ["AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB",
       "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO"]

st.set_page_config(page_title="Votos por escola – TSE", page_icon="🗳️", layout="wide")
st.title("🗳️ Votos por escola e por seção")
st.caption("Fonte: arquivo 'Votação por seção eleitoral' do TSE (dadosabertos.tse.jus.br). "
           "Tudo roda no seu computador.")

with st.sidebar:
    st.header("1. Dados")
    ano = st.selectbox("Ano da eleição", [2026, 2022, 2018], index=0)
    uf = st.selectbox("UF", UFS, index=UFS.index("CE"))
    cargo = st.selectbox("Cargo", tse.CARGOS, index=0)
    turno = st.radio("Turno", [1, 2], horizontal=True)
    origem = st.radio("Origem do arquivo",
                      ["Baixar do TSE", "Caminho de um zip no computador", "Enviar zip"])
    forcar = False
    caminho = ""
    enviado = None
    if origem == "Baixar do TSE":
        forcar = st.checkbox("Baixar de novo (ignorar cópia local)")
    elif origem == "Caminho de um zip no computador":
        caminho = st.text_input("Caminho do .zip")
    else:
        enviado = st.file_uploader("Arquivo .zip", type="zip")
    carregar = st.button("Carregar dados", type="primary")

if carregar:
    barra = st.progress(0.0, text="Preparando…")
    try:
        if origem == "Baixar do TSE":
            barra.progress(0.0, text="Baixando do TSE…")
            zp = tse.baixar_zip(ano, uf, PASTA, forcar, lambda p: barra.progress(p, text="Baixando do TSE…"))
        elif origem == "Caminho de um zip no computador":
            zp = Path(caminho.strip().strip('"'))
            if not zp.is_file():
                raise FileNotFoundError(f"Arquivo não encontrado: {zp}")
        else:
            if enviado is None:
                raise ValueError("Selecione um arquivo .zip.")
            PASTA.mkdir(exist_ok=True)
            zp = PASTA / enviado.name
            zp.write_bytes(enviado.getbuffer())
        barra.progress(0.0, text="Lendo o arquivo…")
        df = tse.carregar(zp, cargo, turno, lambda p: barra.progress(p, text="Lendo o arquivo…"))
        st.session_state["df"] = df
        st.session_state["meta"] = dict(ano=ano, uf=uf, cargo=cargo, turno=turno)
        barra.empty()
    except Exception as e:  # noqa: BLE001
        barra.empty()
        st.sidebar.error(str(e))

if "df" not in st.session_state:
    st.info("Escolha ano, UF e cargo na barra lateral e clique em **Carregar dados**.")
    st.stop()

df = st.session_state["df"]
meta = st.session_state["meta"]
st.success(f"Dados carregados: {meta['cargo']} · {meta['uf']} · {meta['ano']} · {meta['turno']}º turno "
           f"({len(df):,} linhas)".replace(",", "."))

st.header("2. Município e candidatos")
c1, c2 = st.columns([1, 2])
with c1:
    muni = st.selectbox("Município", tse.municipios(df))
    incluir = st.checkbox("Listar brancos e nulos")
cands = tse.candidatos(df, muni, incluir)
mapa = dict(zip(cands["rotulo"], cands["NR_VOTAVEL"]))
votos = dict(zip(cands["rotulo"], cands["QT_VOTOS"]))
with c2:
    escolhidos = st.multiselect("Candidatos (digite nome ou número; ordenados por votos no município)",
                                list(mapa), format_func=lambda r: f"{r} – {votos[r]:,} votos".replace(",", "."))
    extras = st.text_input("Ou digite números de candidato separados por vírgula (ex.: 40114, 12345)")

numeros = [mapa[r] for r in escolhidos]
for n in [x.strip() for x in extras.split(",") if x.strip()]:
    if n in set(cands["NR_VOTAVEL"]):
        if n not in numeros:
            numeros.append(n)
    else:
        st.warning(f"Número {n} não encontrado em {muni}.")

if not numeros:
    st.stop()

t, rotulos = tse.tabela_secoes(df, muni, numeros)
titulo = f"Votos por escola – {muni} – {meta['cargo']} {meta['ano']} ({meta['turno']}º turno)"

st.header("3. Resultado")
esc = tse.tabela_escolas(t, rotulos)
m1, m2, m3 = st.columns(3)
m1.metric("Escolas / locais", len(esc))
m2.metric("Seções", len(t))
m3.metric("Votos (selecionados)", f"{int(t[rotulos].to_numpy().sum()):,}".replace(",", "."))

vis = esc.rename(columns={"NR_ZONA": "Zona", "NR_LOCAL_VOTACAO": "Cód. local",
                          "NM_LOCAL_VOTACAO": "Escola", "DS_LOCAL_VOTACAO_ENDERECO": "Endereço",
                          "N_SECOES": "Seções"})
tab1, tab2 = st.tabs(["Por escola", "Por seção"])
tab1.dataframe(vis, use_container_width=True, hide_index=True)
tab2.dataframe(t.rename(columns={"NR_ZONA": "Zona", "NR_SECAO": "Seção", "NR_LOCAL_VOTACAO": "Cód. local",
                                 "NM_LOCAL_VOTACAO": "Escola", "DS_LOCAL_VOTACAO_ENDERECO": "Endereço"}),
               use_container_width=True, hide_index=True)

st.subheader("Baixar Excel")
nome = f"votos_{tse.slug(muni)}_{tse.slug(meta['cargo'])}_{meta['ano']}"
b1, b2 = st.columns(2)
b1.download_button("⬇️ Uma planilha com todos os selecionados",
                   tse.gerar_excel(t, rotulos, titulo), f"{nome}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
if len(rotulos) > 1:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fn, conteudo in tse.gerar_por_candidato(t, rotulos, titulo).items():
            z.writestr(fn, conteudo)
    b2.download_button("⬇️ Uma planilha por candidato (.zip)", buf.getvalue(), f"{nome}_por_candidato.zip",
                       "application/zip")
st.caption("As somas são fórmulas do Excel; ao abrir o arquivo, ele calcula sozinho.")
