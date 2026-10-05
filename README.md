# Votos por escola (TSE)

Gera planilha Excel com os votos de um ou mais candidatos, por escola e por seção, para qualquer município.

## Como usar
```
pip install -r requirements.txt
streamlit run app.py
```
1. Barra lateral: ano, UF, cargo, turno e origem do arquivo (baixar do TSE, zip local ou upload) → **Carregar dados**.
2. Escolha o município e os candidatos (nome ou número).
3. Baixe o Excel (um com todos, ou um por candidato).

O arquivo usado é `votacao_secao_<ano>_<UF>.zip` (cdn.tse.jus.br). A primeira leitura é demorada; depois fica em cache (`dados/*.parquet`).
Se o TSE ainda não publicou o arquivo do ano, o app avisa; use um zip baixado manualmente.
