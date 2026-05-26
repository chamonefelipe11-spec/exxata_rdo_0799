import re
from io import BytesIO

import pandas as pd
import pdfplumber
import streamlit as st


def extrair_mao_obra_e_equipamentos(arquivos_pdf):
    dados_consolidados = []
    padrao_data = re.compile(r"(\d{2}/\d{2}/\d{4})")

    for arquivo in arquivos_pdf:
        try:
            with pdfplumber.open(arquivo) as pdf:
                texto_primeira_pagina = pdf.pages[0].extract_text() or ""
                data_rdo = "Data não encontrada"

                match_data = padrao_data.search(texto_primeira_pagina)
                if match_data:
                    data_rdo = match_data.group(1)

                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables()

                    for tabela in tabelas:
                        if not tabela or not tabela[0]:
                            continue

                        cabecalho = str(tabela[0][0]).lower() if tabela[0][0] else ""

                        if "mão de obra" in cabecalho or "equipamentos" in cabecalho:
                            for linha in tabela:
                                for celula in linha:
                                    if celula and isinstance(celula, str):
                                        celula_limpa = celula.replace("\n", " ").strip()
                                        celula_limpa = re.sub(r"\s+", " ", celula_limpa)

                                        if (
                                            "mão de obra" in celula_limpa.lower()
                                            or "equipamentos" in celula_limpa.lower()
                                            or celula_limpa == ""
                                        ):
                                            continue

                                        match = re.search(r"^(.*?)\s+(\d+)$", celula_limpa)

                                        if match:
                                            item = match.group(1).strip().upper()
                                            quantidade = match.group(2).zfill(2)

                                            dados_consolidados.append({
                                                "Arquivo": arquivo.name,
                                                "Data": data_rdo,
                                                "Descrição": item,
                                                "Quantidade": quantidade
                                            })

        except Exception as e:
            dados_consolidados.append({
                "Arquivo": arquivo.name,
                "Data": "Erro",
                "Descrição": f"Erro ao processar arquivo: {e}",
                "Quantidade": ""
            })

    return pd.DataFrame(dados_consolidados)


def extrair_atividades(arquivos_pdf):
    dados_extraidos = []

    regex_data = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
    regex_limpa_codigo = re.compile(r"^\d+(?:\.\d+)*\s*[-–]?\s*")

    for arquivo in arquivos_pdf:
        try:
            with pdfplumber.open(arquivo) as pdf:
                data_rdo = "Data não encontrada"

                if len(pdf.pages) > 0:
                    texto_pag1 = pdf.pages[0].extract_text() or ""
                    match_data = regex_data.search(texto_pag1)
                    if match_data:
                        data_rdo = match_data.group(0)

                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables()

                    for tabela in tabelas:
                        for linha in tabela:
                            linha_limpa = [
                                str(celula).replace("\n", " ").strip() if celula else ""
                                for celula in linha
                            ]

                            texto_linha_completa = " ".join(linha_limpa).lower()

                            if "%" in texto_linha_completa and (
                                "concluída" in texto_linha_completa
                                or "andamento" in texto_linha_completa
                            ):
                                atividade_bruta = linha_limpa[0]
                                atividade_limpa = regex_limpa_codigo.sub("", atividade_bruta).strip()

                                if atividade_limpa:
                                    dados_extraidos.append({
                                        "Arquivo": arquivo.name,
                                        "Data": data_rdo,
                                        "Atividade": atividade_limpa
                                    })

        except Exception as e:
            dados_extraidos.append({
                "Arquivo": arquivo.name,
                "Data": "Erro",
                "Atividade": f"Erro ao processar arquivo: {e}"
            })

    return pd.DataFrame(dados_extraidos)


def dataframe_para_excel(df):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")

    buffer.seek(0)
    return buffer


st.set_page_config(
    page_title="Extrator de RDOs",
    layout="wide"
)

st.title("Extrator de dados de RDOs em PDF")

st.write(
    "Faça upload dos PDFs de RDO e escolha o tipo de extração desejada."
)

tipo_extracao = st.radio(
    "Tipo de extração",
    [
        "Mão de obra e equipamentos",
        "Atividades"
    ]
)

arquivos_pdf = st.file_uploader(
    "Selecione os PDFs",
    type=["pdf"],
    accept_multiple_files=True
)

if arquivos_pdf:
    st.success(f"{len(arquivos_pdf)} arquivo(s) carregado(s).")

    if st.button("Processar PDFs"):
        with st.spinner("Processando arquivos..."):
            if tipo_extracao == "Mão de obra e equipamentos":
                df = extrair_mao_obra_e_equipamentos(arquivos_pdf)
                nome_arquivo = "consolidado_mao_obra_e_equipamentos.xlsx"
            else:
                df = extrair_atividades(arquivos_pdf)
                nome_arquivo = "0799_atividades_rdo.xlsx"

        if df.empty:
            st.warning("Nenhum dado foi encontrado nos PDFs enviados.")
        else:
            st.subheader("Prévia dos dados extraídos")
            st.dataframe(df, use_container_width=True)

            excel = dataframe_para_excel(df)

            st.download_button(
                label="Baixar Excel",
                data=excel,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Envie um ou mais PDFs para iniciar.")