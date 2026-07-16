import re
import os
import unicodedata
import difflib
from io import BytesIO
from datetime import datetime, date

import pandas as pd
import pdfplumber
import fitz  # PyMuPDF
import streamlit as st


# ============================================================
# FUNÇÕES AUXILIARES GERAIS
# ============================================================

MESES_PT = {
    "jan": 1,
    "fev": 2,
    "mar": 3,
    "abr": 4,
    "mai": 5,
    "jun": 6,
    "jul": 7,
    "ago": 8,
    "set": 9,
    "out": 10,
    "nov": 11,
    "dez": 12,
}

STOPWORDS = {
    "de", "da", "do", "das", "dos", "e", "em", "para", "por",
    "com", "no", "na", "nos", "nas", "ao", "a", "o", "as", "os"
}


def resetar_arquivo(arquivo):
    try:
        arquivo.seek(0)
    except Exception:
        pass


def obter_nome_arquivo(arquivo):
    return getattr(arquivo, "name", os.path.basename(str(arquivo)))


def limpar_celula(valor):
    if valor is None:
        return ""

    valor = str(valor).replace("\n", " ").strip()
    valor = re.sub(r"\s+", " ", valor)

    return valor


def normalizar_texto(texto):
    texto = limpar_celula(texto).lower()

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))

    texto = texto.replace("ø", "o").replace("Ø", "o")

    texto = re.sub(r"^\d+(?:\.\d+)*\s*[-–]?\s*", "", texto)
    texto = re.sub(r"\bobs\.?.*$", "", texto)

    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    texto = re.sub(r"\s+", " ", texto).strip()

    return texto


def tokens_relevantes(texto):
    tokens = normalizar_texto(texto).split()
    return set(token for token in tokens if token not in STOPWORDS)


def dataframe_para_excel(df):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")

    buffer.seek(0)
    return buffer


def dataframes_para_excel(abas):
    buffer = BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for nome_aba, df in abas.items():
            nome_seguro = re.sub(r"[\[\]\:\*\?\/\\]", "", nome_aba)[:31]
            df.to_excel(writer, index=False, sheet_name=nome_seguro)

    buffer.seek(0)
    return buffer


# ============================================================
# EXTRAÇÃO LEGADA: MÃO DE OBRA E EQUIPAMENTOS
# ============================================================

def extrair_mao_obra_e_equipamentos(arquivos_pdf):
    dados_consolidados = []
    padrao_data = re.compile(r"(\d{2}/\d{2}/\d{4})")

    for arquivo in arquivos_pdf:
        try:
            resetar_arquivo(arquivo)

            with pdfplumber.open(arquivo) as pdf:
                texto_primeira_pagina = pdf.pages[0].extract_text() or ""
                data_rdo = "Data não encontrada"

                match_data = padrao_data.search(texto_primeira_pagina)
                if match_data:
                    data_rdo = match_data.group(1)

                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables() or []

                    for tabela in tabelas:
                        if not tabela or not tabela[0]:
                            continue

                        cabecalho = str(tabela[0][0]).lower() if tabela[0][0] else ""

                        if "mão de obra" in cabecalho or "equipamentos" in cabecalho:
                            for linha in tabela:
                                for celula in linha:
                                    if celula and isinstance(celula, str):
                                        celula_limpa = limpar_celula(celula)

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
                                                "Arquivo": obter_nome_arquivo(arquivo),
                                                "Data": data_rdo,
                                                "Descrição": item,
                                                "Quantidade": quantidade
                                            })

        except Exception as e:
            dados_consolidados.append({
                "Arquivo": obter_nome_arquivo(arquivo),
                "Data": "Erro",
                "Descrição": f"Erro ao processar arquivo: {e}",
                "Quantidade": ""
            })

    return pd.DataFrame(dados_consolidados)


# ============================================================
# EXTRAÇÃO LEGADA: ATIVIDADES DO RDO
# ============================================================

def extrair_atividades(arquivos_pdf):
    dados_extraidos = []

    regex_data = re.compile(r"\b\d{2}/\d{2}/\d{4}\b")
    regex_limpa_codigo = re.compile(r"^\d+(?:\.\d+)*\s*[-–]?\s*")

    for arquivo in arquivos_pdf:
        try:
            resetar_arquivo(arquivo)

            with pdfplumber.open(arquivo) as pdf:
                data_rdo = "Data não encontrada"

                if len(pdf.pages) > 0:
                    texto_pag1 = pdf.pages[0].extract_text() or ""
                    match_data = regex_data.search(texto_pag1)
                    if match_data:
                        data_rdo = match_data.group(0)

                for pagina in pdf.pages:
                    tabelas = pagina.extract_tables() or []

                    for tabela in tabelas:
                        for linha in tabela:
                            linha_limpa = [
                                limpar_celula(celula)
                                for celula in linha
                            ]

                            texto_linha_completa = " ".join(linha_limpa).lower()

                            if "%" in texto_linha_completa and (
                                "concluída" in texto_linha_completa
                                or "concluida" in texto_linha_completa
                                or "andamento" in texto_linha_completa
                            ):
                                atividade_bruta = linha_limpa[0]
                                atividade_limpa = regex_limpa_codigo.sub("", atividade_bruta).strip()

                                if atividade_limpa:
                                    dados_extraidos.append({
                                        "Arquivo": obter_nome_arquivo(arquivo),
                                        "Data": data_rdo,
                                        "Atividade": atividade_limpa
                                    })

        except Exception as e:
            dados_extraidos.append({
                "Arquivo": obter_nome_arquivo(arquivo),
                "Data": "Erro",
                "Atividade": f"Erro ao processar arquivo: {e}"
            })

    return pd.DataFrame(dados_extraidos)


# ============================================================
# EXTRAÇÃO: COMENTÁRIOS DO RDO
# ============================================================

RE_DATA_RDO_COMENTARIOS = [
    re.compile(r"Data\s+do\s+relat[oó]rio\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})", re.I),
    re.compile(r"Relat[oó]rio\s+(\d{2}/\d{2}/\d{4})\s+n[°ºo]?", re.I),
]

RE_INICIO_COMENTARIOS = re.compile(
    r"^\s*Coment[aá]rios\s*\(\s*\d+\s*\)\s*$",
    re.I | re.M,
)

RE_FIM_COMENTARIOS = re.compile(
    r"^\s*(?:Fotos|Assinaturas|Anexos|Observa[cç][oõ]es|Ocorr[eê]ncias)\s*(?:\(\s*\d+\s*\))?\s*$",
    re.I | re.M,
)

RE_METADADO_AUTOR_COMENTARIO = re.compile(
    r"^\s*[^\n]{1,80}?\s+\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s*$",
    re.M,
)

RE_NOTA_COMENTARIO = re.compile(
    r"(?ms)^\s*(Nota\s*\d+\s*:\s*.*?)(?=^\s*Nota\s*\d+\s*:|\Z)",
    re.I,
)


def normalizar_texto_comentarios(texto):
    texto = texto.replace("\u00a0", " ").replace("\r", "\n")
    texto = re.sub(r"(?<=\w)-\n(?=\w)", "", texto)
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


def extrair_texto_pdf_comentarios(arquivo):
    resetar_arquivo(arquivo)
    conteudo = arquivo.read()
    resetar_arquivo(arquivo)

    partes = []
    with fitz.open(stream=conteudo, filetype="pdf") as pdf:
        for pagina in pdf:
            partes.append(pagina.get_text("text", sort=True) or "")

    return normalizar_texto_comentarios("\n".join(partes))


def extrair_data_rdo_comentarios(texto, nome_arquivo):
    for padrao in RE_DATA_RDO_COMENTARIOS:
        achado = padrao.search(texto)
        if achado:
            return achado.group(1)

    achado = re.search(r"(\d{2})[-_](\d{2})[-_](\d{4})", nome_arquivo)
    if achado:
        return f"{achado.group(1)}/{achado.group(2)}/{achado.group(3)}"

    return "Data não encontrada"


def extrair_secao_comentarios(texto):
    inicio = RE_INICIO_COMENTARIOS.search(texto)
    if not inicio:
        return ""

    restante = texto[inicio.end():]
    fim = RE_FIM_COMENTARIOS.search(restante)
    if fim:
        restante = restante[:fim.start()]

    restante = re.sub(
        r"^\s*Relat[oó]rio\s+\d{2}/\d{2}/\d{4}\s+n[°ºo]?\s*\d+\s*$",
        "",
        restante,
        flags=re.I | re.M,
    )
    restante = re.sub(r"^\s*\d+\s*/\s*\d+\s*$", "", restante, flags=re.M)
    return restante.strip()


def limpar_comentario_rdo(texto):
    texto = RE_METADADO_AUTOR_COMENTARIO.sub("", texto)
    texto = re.sub(r"\s*\n\s*", " ", texto)
    texto = re.sub(r"\s{2,}", " ", texto)
    return texto.strip(" -\n\t")


def extrair_lista_comentarios(texto):
    secao = extrair_secao_comentarios(texto)
    if not secao:
        return []

    notas = [
        limpar_comentario_rdo(match.group(1))
        for match in RE_NOTA_COMENTARIO.finditer(secao)
    ]
    notas = [nota for nota in notas if nota]
    if notas:
        return notas

    blocos = RE_METADADO_AUTOR_COMENTARIO.split(secao)
    comentarios = [limpar_comentario_rdo(bloco) for bloco in blocos]
    return [comentario for comentario in comentarios if comentario]


def extrair_comentarios_rdo(arquivos_pdf):
    dados = []

    for arquivo in arquivos_pdf:
        nome_arquivo = obter_nome_arquivo(arquivo)

        try:
            texto = extrair_texto_pdf_comentarios(arquivo)
            data_rdo = extrair_data_rdo_comentarios(texto, nome_arquivo)
            comentarios = extrair_lista_comentarios(texto)

            for comentario in comentarios:
                dados.append({
                    "Data do RDO": data_rdo,
                    "Comentário": comentario,
                    "Nome do arquivo": nome_arquivo,
                })

        except Exception as e:
            dados.append({
                "Data do RDO": "Erro",
                "Comentário": f"Erro ao processar arquivo: {e}",
                "Nome do arquivo": nome_arquivo,
            })

    return pd.DataFrame(
        dados,
        columns=["Data do RDO", "Comentário", "Nome do arquivo"]
    )


# ============================================================
# COMPARATIVO: PROGRAMAÇÃO SEMANAL X RDO
# ============================================================

def extrair_info_rdo(arquivo_rdo):
    resetar_arquivo(arquivo_rdo)

    data_rdo = "Data não encontrada"
    numero_rdo = "Nº não encontrado"
    texto_total = ""

    with pdfplumber.open(arquivo_rdo) as pdf:
        for indice_pagina, pagina in enumerate(pdf.pages):
            texto_pagina = pagina.extract_text() or ""
            texto_total += texto_pagina + "\n"

            if indice_pagina == 0:
                match_data = re.search(r"\b\d{2}/\d{2}/\d{4}\b", texto_pagina)
                if match_data:
                    data_rdo = match_data.group(0)

                match_numero = re.search(
                    r"Relat[óo]rio(?:\s+\d{2}/\d{2}/\d{4})?\s+n[°º]?\s*(\d+)",
                    texto_pagina,
                    flags=re.IGNORECASE
                )

                if not match_numero:
                    match_numero = re.search(
                        r"\bn[°º]\s*(\d+)\b",
                        texto_pagina,
                        flags=re.IGNORECASE
                    )

                if match_numero:
                    numero_rdo = match_numero.group(1)

    return data_rdo, numero_rdo, texto_total


def extrair_atividades_rdo_comparativo(arquivo_rdo):
    data_rdo, numero_rdo, texto_total = extrair_info_rdo(arquivo_rdo)

    atividades = []
    sem_atividades = "sem atividades" in texto_total.lower()

    resetar_arquivo(arquivo_rdo)

    with pdfplumber.open(arquivo_rdo) as pdf:
        for pagina in pdf.pages:
            tabelas = pagina.extract_tables() or []

            for tabela in tabelas:
                if not tabela:
                    continue

                for linha in tabela:
                    linha_limpa = [limpar_celula(celula) for celula in linha]

                    if not linha_limpa:
                        continue

                    primeira_coluna = linha_limpa[0]
                    texto_linha = " ".join(linha_limpa).lower()

                    if primeira_coluna.lower().startswith("atividades"):
                        continue

                    if "%" in texto_linha and (
                        "andamento" in texto_linha
                        or "conclu" in texto_linha
                    ):
                        atividade = re.sub(
                            r"^\d+(?:\.\d+)*\s*[-–]\s*",
                            "",
                            primeira_coluna
                        ).strip()

                        if atividade:
                            atividades.append(atividade)

    atividades = list(dict.fromkeys(atividades))

    return {
        "arquivo": obter_nome_arquivo(arquivo_rdo),
        "data": data_rdo,
        "numero": numero_rdo,
        "atividades": atividades,
        "texto_total": texto_total,
        "sem_atividades": sem_atividades
    }


def converter_data_programacao(valor, ano_referencia):
    valor = limpar_celula(valor).lower()

    match = re.search(r"\b(\d{1,2})\s*[-/]\s*([a-zç]{3})\b", valor)
    if not match:
        return None

    dia = int(match.group(1))
    mes_texto = match.group(2)[:3]
    mes = MESES_PT.get(mes_texto)

    if not mes:
        return None

    return date(ano_referencia, mes, dia)


def categoria_base(categoria):
    categoria = limpar_celula(categoria).upper()
    categoria = re.sub(r"^\d+\s+", "", categoria).strip()

    if "OBRAS CIVIL" in categoria or "OBRA CIVIL" in categoria:
        return "OBRAS CIVIL"

    if "ELÉTRICA" in categoria or "ELETRICA" in categoria:
        return "ELÉTRICA"

    if "MONTAGEM" in categoria:
        return "MONTAGEM"

    if "TERCEIRO" in categoria:
        return "TERCEIRO"

    if "COMISSIONAMENTO" in categoria:
        return "COMISSIONAMENTO"

    return categoria.split("-")[0].strip()


def eh_linha_categoria(linha):
    if not linha:
        return False

    primeira_coluna = limpar_celula(linha[0])

    if not primeira_coluna:
        return False

    if primeira_coluna.upper() == "ID":
        return False

    if len(linha) > 1 and limpar_celula(linha[1]):
        return False

    texto = primeira_coluna.upper()

    categorias_chave = [
        "OBRAS CIVIL",
        "OBRA CIVIL",
        "ELÉTRICA",
        "ELETRICA",
        "MONTAGEM",
        "TERCEIRO",
        "COMISSIONAMENTO"
    ]

    return any(chave in texto for chave in categorias_chave)


def extrair_local_programacao(arquivo_programacao):
    resetar_arquivo(arquivo_programacao)

    local = "SE MANAUS"

    try:
        with pdfplumber.open(arquivo_programacao) as pdf:
            if len(pdf.pages) > 0:
                texto = pdf.pages[0].extract_text() or ""

                match = re.search(
                    r"PROGRAMAÇÃO SEMANAL DE ATIVIDADES[-–]\s*([A-ZÁÉÍÓÚÃÕÇ\s]+)",
                    texto,
                    flags=re.IGNORECASE
                )

                if match:
                    local = match.group(1)
                    local = local.split("Categoria")[0].strip().upper()

    except Exception:
        local = "SE MANAUS"

    return local


def extrair_programadas_por_data(arquivo_programacao, data_referencia):
    dados = []
    ano_referencia = data_referencia.year

    resetar_arquivo(arquivo_programacao)

    with pdfplumber.open(arquivo_programacao) as pdf:
        categoria_atual = ""

        for numero_pagina, pagina in enumerate(pdf.pages, start=1):
            tabelas = pagina.extract_tables() or []

            for tabela in tabelas:
                if not tabela:
                    continue

                mapa_datas = {}

                for linha in tabela:
                    for indice_coluna, celula in enumerate(linha):
                        data_coluna = converter_data_programacao(
                            celula,
                            ano_referencia
                        )

                        if data_coluna:
                            mapa_datas[indice_coluna] = data_coluna

                colunas_data = [
                    coluna for coluna, data_coluna in mapa_datas.items()
                    if data_coluna == data_referencia
                ]

                if not colunas_data:
                    continue

                coluna_data = colunas_data[0]

                for indice_linha, linha in enumerate(tabela):
                    linha_limpa = [limpar_celula(celula) for celula in linha]

                    if not linha_limpa:
                        continue

                    if eh_linha_categoria(linha_limpa):
                        categoria_atual = categoria_base(linha_limpa[0])
                        continue

                    if len(linha_limpa) <= coluna_data:
                        continue

                    id_item = linha_limpa[0] if len(linha_limpa) > 0 else ""
                    atividade = linha_limpa[1] if len(linha_limpa) > 1 else ""
                    status_linha = linha_limpa[2] if len(linha_limpa) > 2 else ""
                    marcador_pr = linha_limpa[3].upper() if len(linha_limpa) > 3 else ""
                    marcado_no_dia = linha_limpa[coluna_data].lower() == "x"

                    proxima_linha = []
                    if indice_linha + 1 < len(tabela):
                        proxima_linha = [
                            limpar_celula(celula)
                            for celula in tabela[indice_linha + 1]
                        ]

                    proxima_eh_r = (
                        len(proxima_linha) > 3
                        and proxima_linha[3].upper() == "R"
                    )

                    # Linha P normal ou linha P com falha de leitura na coluna P/R
                    parece_linha_p = (
                        marcador_pr == "P"
                        or (
                            id_item
                            and atividade
                            and id_item.upper() != "ID"
                            and marcador_pr == ""
                            and status_linha == ""
                            and marcado_no_dia
                        )
                    )

                    if parece_linha_p and atividade and marcado_no_dia:
                        status_programacao = ""

                        if proxima_linha and len(proxima_linha) > 2:
                            status_programacao = proxima_linha[2]

                        descricao_servico = ""
                        if len(linha_limpa) > 19:
                            descricao_servico = linha_limpa[19]

                        dados.append({
                            "Data": data_referencia.strftime("%d/%m/%Y"),
                            "Categoria": categoria_atual,
                            "ID": id_item,
                            "Atividade programada": atividade,
                            "Status programação": status_programacao,
                            "Descrição serviço": descricao_servico,
                            "Página programação": numero_pagina
                        })

    return pd.DataFrame(dados)


def possui_conflito_tecnico(texto_a, texto_b):
    """
    Evita falsos positivos entre atividades tecnicamente semelhantes,
    mas que não devem ser consideradas iguais.
    Exemplo:
    - SEC c/ LT x SEC s/ LT
    - TC x TP
    - DJ x TC/TP
    """

    norm_a = normalizar_texto(texto_a)
    norm_b = normalizar_texto(texto_b)

    a_com_lt = bool(re.search(r"\bc\s*lt\b", norm_a))
    a_sem_lt = bool(re.search(r"\bs\s*lt\b", norm_a))
    b_com_lt = bool(re.search(r"\bc\s*lt\b", norm_b))
    b_sem_lt = bool(re.search(r"\bs\s*lt\b", norm_b))

    if (a_com_lt and b_sem_lt) or (a_sem_lt and b_com_lt):
        return True

    tokens_a = tokens_relevantes(texto_a)
    tokens_b = tokens_relevantes(texto_b)

    pares_conflitantes = [
        ("tc", "tp"),
        ("dj", "tc"),
        ("dj", "tp"),
    ]

    for item_a, item_b in pares_conflitantes:
        conflito_ab = (
            item_a in tokens_a
            and item_b not in tokens_a
            and item_b in tokens_b
            and item_a not in tokens_b
        )

        conflito_ba = (
            item_b in tokens_a
            and item_a not in tokens_a
            and item_a in tokens_b
            and item_b not in tokens_b
        )

        if conflito_ab or conflito_ba:
            return True

    return False


def calcular_score_correspondencia(atividade_programada, atividade_rdo):
    if possui_conflito_tecnico(atividade_programada, atividade_rdo):
        return 0.0

    prog_norm = normalizar_texto(atividade_programada)
    rdo_norm = normalizar_texto(atividade_rdo)

    if not prog_norm or not rdo_norm:
        return 0.0

    if prog_norm == rdo_norm:
        return 1.0

    tokens_prog = tokens_relevantes(atividade_programada)
    tokens_rdo = tokens_relevantes(atividade_rdo)

    if not tokens_prog or not tokens_rdo:
        return 0.0

    intersecao = tokens_prog.intersection(tokens_rdo)
    score_tokens = len(intersecao) / max(len(tokens_prog), len(tokens_rdo))

    score_difflib = difflib.SequenceMatcher(
        None,
        prog_norm,
        rdo_norm
    ).ratio()

    # Usa a média ponderada, privilegiando tokens relevantes.
    score_final = (score_tokens * 0.75) + (score_difflib * 0.25)

    return score_final


def melhor_correspondencia(atividade_programada, atividades_rdo, limite_correspondencia):
    melhor_score = 0.0
    melhor_atividade = ""

    for atividade_rdo in atividades_rdo:
        score = calcular_score_correspondencia(
            atividade_programada,
            atividade_rdo
        )

        if score > melhor_score:
            melhor_score = score
            melhor_atividade = atividade_rdo

    encontrada = melhor_score >= limite_correspondencia

    return encontrada, melhor_atividade, melhor_score


def comparar_programacao_x_rdo(
    arquivo_programacao,
    arquivos_rdo,
    limite_correspondencia=0.75,
    categorias_ignorar=None
):
    if categorias_ignorar is None:
        categorias_ignorar = []

    resultados = []

    for arquivo_rdo in arquivos_rdo:
        info_rdo = extrair_atividades_rdo_comparativo(arquivo_rdo)

        if info_rdo["data"] == "Data não encontrada":
            resultados.append({
                "Arquivo RDO": obter_nome_arquivo(arquivo_rdo),
                "Nº RDO": info_rdo["numero"],
                "Data": info_rdo["data"],
                "Categoria": "",
                "ID": "",
                "Atividade programada": "Data do RDO não encontrada",
                "Encontrada no RDO?": "Não",
                "Atividade RDO mais próxima": "",
                "Score correspondência": 0,
                "Status programação": "",
                "Descrição serviço": "",
                "Página programação": ""
            })
            continue

        data_referencia = datetime.strptime(
            info_rdo["data"],
            "%d/%m/%Y"
        ).date()

        df_programadas = extrair_programadas_por_data(
            arquivo_programacao,
            data_referencia
        )

        if not df_programadas.empty and categorias_ignorar:
            df_programadas = df_programadas[
                ~df_programadas["Categoria"].isin(categorias_ignorar)
            ]

        if df_programadas.empty:
            resultados.append({
                "Arquivo RDO": obter_nome_arquivo(arquivo_rdo),
                "Nº RDO": info_rdo["numero"],
                "Data": info_rdo["data"],
                "Categoria": "",
                "ID": "",
                "Atividade programada": "Nenhuma atividade programada encontrada para a data",
                "Encontrada no RDO?": "Não",
                "Atividade RDO mais próxima": "",
                "Score correspondência": 0,
                "Status programação": "",
                "Descrição serviço": "",
                "Página programação": ""
            })
            continue

        for _, linha in df_programadas.iterrows():
            if info_rdo["sem_atividades"]:
                encontrada = False
                melhor_atividade = "RDO descrito como Sem atividades"
                score = 0.0
            else:
                encontrada, melhor_atividade, score = melhor_correspondencia(
                    linha["Atividade programada"],
                    info_rdo["atividades"],
                    limite_correspondencia
                )

            resultados.append({
                "Arquivo RDO": obter_nome_arquivo(arquivo_rdo),
                "Nº RDO": info_rdo["numero"],
                "Data": info_rdo["data"],
                "Categoria": linha["Categoria"],
                "ID": linha["ID"],
                "Atividade programada": linha["Atividade programada"],
                "Encontrada no RDO?": "Sim" if encontrada else "Não",
                "Atividade RDO mais próxima": melhor_atividade,
                "Score correspondência": round(score, 3),
                "Status programação": linha["Status programação"],
                "Descrição serviço": linha["Descrição serviço"],
                "Página programação": linha["Página programação"]
            })

    return pd.DataFrame(resultados)


def gerar_mensagem_pendencias(
    df_comparativo_rdo,
    local="SE MANAUS",
    incluir_categorias_vazias=True
):
    if df_comparativo_rdo.empty:
        return ""

    numero_rdo = str(df_comparativo_rdo["Nº RDO"].iloc[0])
    data_rdo = str(df_comparativo_rdo["Data"].iloc[0])

    df_pendencias = df_comparativo_rdo[
        df_comparativo_rdo["Encontrada no RDO?"] == "Não"
    ].copy()

    df_pendencias = df_pendencias[
        df_pendencias["Categoria"].astype(str).str.strip() != ""
    ]

    linhas = [
        f"*{local.upper()}*",
        f"Nº {numero_rdo} – {data_rdo}",
    ]

    if df_pendencias.empty:
        linhas.append("Não foram identificadas atividades programadas ausentes no RDO.")
        return "\n".join(linhas)

    linhas.extend([
        'As atividades listadas abaixo se encontram na programação semanal, mas não se encontram no RDO ou foram descritas “Sem atividades”:',
        ""
    ])

    categorias_padrao = [
        "OBRAS CIVIL",
        "ELÉTRICA",
        "MONTAGEM",
        "TERCEIRO",
        "COMISSIONAMENTO"
    ]

    outras_categorias = [
        categoria for categoria in df_pendencias["Categoria"].dropna().unique()
        if categoria and categoria not in categorias_padrao
    ]

    categorias = categorias_padrao + list(outras_categorias)

    for categoria in categorias:
        df_categoria = df_pendencias[df_pendencias["Categoria"] == categoria]

        if df_categoria.empty and not incluir_categorias_vazias:
            continue

        linhas.append(f"> {categoria}")

        for _, linha in df_categoria.iterrows():
            linhas.append(f"- {linha['Atividade programada']}")

        linhas.append("")

    return "\n".join(linhas).strip()


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

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
        "Atividades",
        "Comentários do RDO",
        "Comparativo programação semanal x RDO"
    ]
)


# ============================================================
# INTERFACE: COMPARATIVO PROGRAMAÇÃO X RDO
# ============================================================

if tipo_extracao == "Comparativo programação semanal x RDO":
    arquivo_programacao = st.file_uploader(
        "Selecione o PDF da programação semanal",
        type=["pdf"],
        accept_multiple_files=False,
        key="programacao_pdf"
    )

    arquivos_rdo = st.file_uploader(
        "Selecione um ou mais PDFs de RDO",
        type=["pdf"],
        accept_multiple_files=True,
        key="rdo_pdf_comparativo"
    )

    col1, col2 = st.columns(2)

    with col1:
        limite_correspondencia = st.slider(
            "Score mínimo para considerar atividade encontrada",
            min_value=0.50,
            max_value=1.00,
            value=0.75,
            step=0.01
        )

    with col2:
        incluir_categorias_vazias = st.checkbox(
            "Incluir categorias vazias na mensagem",
            value=True
        )

    categorias_ignorar = st.multiselect(
        "Categorias a ignorar no comparativo, se aplicável",
        options=[
            "OBRAS CIVIL",
            "ELÉTRICA",
            "MONTAGEM",
            "TERCEIRO",
            "COMISSIONAMENTO"
        ],
        default=[]
    )

    if arquivo_programacao and arquivos_rdo:
        st.success(
            f"Programação semanal carregada e {len(arquivos_rdo)} RDO(s) carregado(s)."
        )

        if st.button("Comparar programação semanal x RDO"):
            with st.spinner("Comparando arquivos..."):
                local = extrair_local_programacao(arquivo_programacao)

                df_comparativo = comparar_programacao_x_rdo(
                    arquivo_programacao=arquivo_programacao,
                    arquivos_rdo=arquivos_rdo,
                    limite_correspondencia=limite_correspondencia,
                    categorias_ignorar=categorias_ignorar
                )

            if df_comparativo.empty:
                st.warning("Nenhum dado foi encontrado para comparação.")
            else:
                df_pendencias = df_comparativo[
                    df_comparativo["Encontrada no RDO?"] == "Não"
                ].copy()

                df_resumo = (
                    df_comparativo
                    .groupby(["Arquivo RDO", "Nº RDO", "Data", "Categoria", "Encontrada no RDO?"], dropna=False)
                    .size()
                    .reset_index(name="Quantidade")
                )

                st.subheader("Comparativo completo")
                st.dataframe(df_comparativo, use_container_width=True)

                st.subheader("Atividades programadas não encontradas no RDO")
                st.dataframe(df_pendencias, use_container_width=True)

                st.subheader("Mensagens sugeridas")

                mensagens_geradas = []

                grupos = df_comparativo.groupby(
                    ["Arquivo RDO", "Nº RDO", "Data"],
                    dropna=False
                )

                for (arquivo_rdo, numero_rdo, data_rdo), df_grupo in grupos:
                    mensagem = gerar_mensagem_pendencias(
                        df_grupo,
                        local=local,
                        incluir_categorias_vazias=incluir_categorias_vazias
                    )

                    mensagens_geradas.append(mensagem)

                    with st.expander(f"Mensagem - RDO Nº {numero_rdo} - {data_rdo}", expanded=True):
                        st.code(mensagem, language="markdown")

                texto_todas_mensagens = "\n\n---\n\n".join(mensagens_geradas)

                st.download_button(
                    label="Baixar mensagens em TXT",
                    data=texto_todas_mensagens.encode("utf-8"),
                    file_name="0799_mensagens_pendencias_rdo.txt",
                    mime="text/plain"
                )

                excel = dataframes_para_excel({
                    "Comparativo": df_comparativo,
                    "Pendencias": df_pendencias,
                    "Resumo": df_resumo
                })

                st.download_button(
                    label="Baixar Excel do comparativo",
                    data=excel,
                    file_name="0799_comparativo_programacao_x_rdo.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    else:
        st.info("Envie a programação semanal e ao menos um RDO para iniciar.")


# ============================================================
# INTERFACE: EXTRAÇÕES LEGADAS
# ============================================================

else:
    arquivos_pdf = st.file_uploader(
        "Selecione os PDFs",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdfs_legado"
    )

    if arquivos_pdf:
        st.success(f"{len(arquivos_pdf)} arquivo(s) carregado(s).")

        if st.button("Processar PDFs"):
            with st.spinner("Processando arquivos..."):
                if tipo_extracao == "Mão de obra e equipamentos":
                    df = extrair_mao_obra_e_equipamentos(arquivos_pdf)
                    nome_arquivo = "consolidado_mao_obra_e_equipamentos.xlsx"
                elif tipo_extracao == "Atividades":
                    df = extrair_atividades(arquivos_pdf)
                    nome_arquivo = "0799_atividades_rdo.xlsx"
                else:
                    df = extrair_comentarios_rdo(arquivos_pdf)
                    nome_arquivo = "0799_comentarios_rdo.xlsx"

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
