import streamlit as st
import pandas as pd
import fitz
import re
from io import BytesIO

st.warning("TESTE NOVO - CLASSES - PORTA 8510")

# ============================================================
# CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="QUIXABA - Analisador P Trab",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def moeda(valor):
    try:
        return (
            f"R$ {float(valor):,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )
    except:
        return "-"


def numero_br(texto):
    """
    Converte números brasileiros:
    25.214,40 -> 25214.40
    7,08 -> 7.08
    """
    if texto is None:
        return None

    texto = str(texto).strip()
    texto = texto.replace("R$", "").replace(" ", "")

    if not texto:
        return None

    try:
        if "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        return float(texto)
    except:
        return None


def limpar_texto(texto):
    if not texto:
        return ""

    texto = texto.replace("\x00", " ")
    texto = re.sub(r"[ \t]+", " ", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)

    return texto.strip()


# ============================================================
# LEITURA DE PDF
# ============================================================

def ler_pdf(arquivo):
    dados = arquivo.getvalue()
    documento = fitz.open(stream=dados, filetype="pdf")

    paginas = []
    texto_total = ""

    for numero_pagina, pagina in enumerate(documento, start=1):
        texto = pagina.get_text("text")
        texto = limpar_texto(texto)

        paginas.append({
            "pagina": numero_pagina,
            "texto": texto
        })

        texto_total += f"\n\n--- PÁGINA {numero_pagina} ---\n{texto}"

    documento.close()

    return limpar_texto(texto_total), paginas


# ============================================================
# LEITURA DE EXCEL
# ============================================================

def ler_excel(arquivo):
    excel = pd.ExcelFile(arquivo)
    abas = {}

    texto_total = ""

    for aba in excel.sheet_names:
        try:
            df = pd.read_excel(
                arquivo,
                sheet_name=aba,
                header=None,
                dtype=object
            )

            abas[aba] = df

            texto_total += f"\n\n--- ABA: {aba} ---\n"

            for _, linha in df.iterrows():
                valores = []

                for valor in linha:
                    if pd.notna(valor):
                        valores.append(str(valor))

                if valores:
                    texto_total += " | ".join(valores) + "\n"

        except Exception as erro:
            abas[aba] = pd.DataFrame(
                {"Erro": [str(erro)]}
            )

    return limpar_texto(texto_total), abas


# ============================================================
# IDENTIFICAÇÃO DO TIPO DO P TRAB
# ============================================================

def identificar_fase(texto):
    texto_upper = texto.upper()

    preparo = (
        "APÊNDICE 2" in texto_upper
        or "APENDICE 2" in texto_upper
        or "PREPARO" in texto_upper
    )

    emprego = (
        "APÊNDICE 4" in texto_upper
        or "APENDICE 4" in texto_upper
        or "EMPREGO" in texto_upper
    )

    if preparo and emprego:
        return "Preparo e Emprego"

    if preparo:
        return "Preparo"

    if emprego:
        return "Emprego"

    return "Não identificado"


# ============================================================
# IDENTIFICAÇÃO DAS CLASSES
# ============================================================

def identificar_classes(texto):
    t = texto.upper()

    classes = []

    padroes = {
        "CLASSE I": [
            r"CLASSE\s*I\b",
            r"ALIMENTA[ÇC][ÃA]O",
            r"\bQR\b",
            r"\bQS\b",
            r"RANCHO"
        ],

        "CLASSE III": [
            r"CLASSE\s*III\b",
            r"COMBUST[IÍ]VEL",
            r"[ÓO]LEO DIESEL",
            r"\bDIESEL\b",
            r"GASOLINA"
        ],

        "CLASSE VI": [
            r"CLASSE\s*VI\b",
            r"MATERIAL DE ENGENHARIA",
            r"MANUTEN[ÇC][ÃA]O DE GERADOR"
        ],

        "CLASSE IX": [
            r"CLASSE\s*IX\b",
            r"MANUTEN[ÇC][ÃA]O DE VIATURA",
            r"MANUTEN[ÇC][ÃA]O DE VIATURAS"
        ]
    }

    for classe, lista_padroes in padroes.items():
        if any(re.search(p, t) for p in lista_padroes):
            classes.append(classe)

    return classes


# ============================================================
# EFETIVO
# ============================================================

def extrair_efetivos(texto):
    padroes = [
        r"EFETIVO(?:\s+EMPREGADO)?\s*[:\-]?\s*(\d{1,5})",
        r"PARA\s+UM\s+EFETIVO\s+DE\s+(\d{1,5})",
        r"(\d{1,5})\s+MILITARES",
        r"EFETIVO\s+DE\s+(\d{1,5})"
    ]

    encontrados = []

    for padrao in padroes:
        resultados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        for resultado in resultados:
            try:
                valor = int(resultado)

                if 0 < valor < 100000:
                    encontrados.append(valor)
            except:
                pass

    return list(dict.fromkeys(encontrados))


# ============================================================
# DIAS
# ============================================================

def extrair_dias(texto):
    padroes = [
        r"N[º°]?\s*DIAS?\s*[:\-]?\s*(\d{1,3})",
        r"PER[IÍ]ODO\s+DE\s+(\d{1,3})\s+DIAS",
        r"POR\s+UM\s+PER[IÍ]ODO\s+DE\s+(\d{1,3})\s+DIAS",
        r"(\d{1,3})\s+DIAS"
    ]

    encontrados = []

    for padrao in padroes:
        resultados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        for resultado in resultados:
            try:
                valor = int(resultado)

                if 0 < valor <= 366:
                    encontrados.append(valor)
            except:
                pass

    return list(dict.fromkeys(encontrados))


# ============================================================
# VALORES MONETÁRIOS
# ============================================================

def extrair_valores(texto):
    padrao = r"R\$\s*([\d\.]+,\d{2})"

    resultados = re.findall(padrao, texto)

    valores = []

    for item in resultados:
        valor = numero_br(item)

        if valor is not None:
            valores.append(valor)

    return valores


# ============================================================
# NATUREZA DA DESPESA
# ============================================================

def extrair_nd(texto):
    padrao = r"\b33\.90\.\d{2}\b"

    resultados = re.findall(padrao, texto)

    return list(dict.fromkeys(resultados))


# ============================================================
# CODUG
# ============================================================

def extrair_codug(texto):
    padroes = [
        r"CODUG\s*[:\-]?\s*(\d{5,6})",
        r"UGE\s*[:\-]?\s*(\d{5,6})"
    ]

    resultados = []

    for padrao in padroes:
        achados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        resultados.extend(achados)

    return list(dict.fromkeys(resultados))


# ============================================================
# COMBUSTÍVEL
# ============================================================

def extrair_combustivel(texto):
    resultados = []

    padroes_litros = [
        r"([\d\.]+,\d+)\s*(?:L|LITROS?)\b",
        r"([\d\.]+)\s*(?:L|LITROS?)\b"
    ]

    litros = []

    for padrao in padroes_litros:
        achados = re.findall(
            padrao,
            texto,
            flags=re.IGNORECASE
        )

        for item in achados:
            valor = numero_br(item)

            if valor is not None and valor > 0:
                litros.append(valor)

    litros = list(dict.fromkeys(litros))

    diesel = bool(
        re.search(
            r"DIESEL|ÓLEO DIESEL|OLEO DIESEL",
            texto,
            re.IGNORECASE
        )
    )

    gasolina = bool(
        re.search(
            r"GASOLINA",
            texto,
            re.IGNORECASE
        )
    )

    if diesel:
        resultados.append("Óleo Diesel")

    if gasolina:
        resultados.append("Gasolina")

    return resultados, litros


# ============================================================
# QR / QS
# ============================================================

def identificar_alimentacao(texto):
    t = texto.upper()

    qr = bool(
        re.search(r"\bQR\b|QUANTITATIVO DE RANCHO \(QR\)", t)
    )

    qs = bool(
        re.search(r"\bQS\b|QUANTITATIVO DE SUBSIST[ÊE]NCIA", t)
    )

    itens = []

    if qr:
        itens.append("QR")

    if qs:
        itens.append("QS")

    return itens


# ============================================================
# EXTRAIR TRECHOS DE CADA CLASSE
# ============================================================

def trecho_classe(texto, classe):
    linhas = texto.splitlines()
    selecionadas = []

    palavras = {
        "CLASSE I": [
            "CLASSE I",
            "ALIMENTAÇÃO",
            "ALIMENTACAO",
            "RANCHO",
            " QR",
            " QS"
        ],

        "CLASSE III": [
            "CLASSE III",
            "COMBUSTÍVEL",
            "COMBUSTIVEL",
            "DIESEL",
            "GASOLINA"
        ],

        "CLASSE VI": [
            "CLASSE VI",
            "GERADOR"
        ],

        "CLASSE IX": [
            "CLASSE IX",
            "VIATURA"
        ]
    }

    termos = palavras.get(classe, [classe])

    for i, linha in enumerate(linhas):
        linha_upper = linha.upper()

        if any(termo in linha_upper for termo in termos):
            inicio = max(0, i - 2)
            fim = min(len(linhas), i + 4)

            bloco = "\n".join(linhas[inicio:fim])

            if bloco not in selecionadas:
                selecionadas.append(bloco)

    return "\n\n".join(selecionadas)


# ============================================================
# ANÁLISE GERAL
# ============================================================

def analisar(texto):
    return {
        "fase": identificar_fase(texto),
        "classes": identificar_classes(texto),
        "efetivos": extrair_efetivos(texto),
        "dias": extrair_dias(texto),
        "valores": extrair_valores(texto),
        "nd": extrair_nd(texto),
        "codug": extrair_codug(texto),
        "alimentacao": identificar_alimentacao(texto),
        "combustiveis": extrair_combustivel(texto)
    }


# ============================================================
# CABEÇALHO
# ============================================================

st.title("📊 QUIXABA")

st.subheader("Sistema de Análise de Planos de Trabalho")

st.write(
    "Leitura e análise automática de **P Trab**, "
    "com identificação de classes, dados financeiros "
    "e memória de cálculo."
)

st.info(
    "O sistema aceita arquivos PDF e Excel (.xlsx)."
)

# ============================================================
# CONFIGURAÇÃO DOS VALORES
# ============================================================

with st.sidebar:
    st.header("⚙️ Valores de referência")

    st.caption(
        "Informe os valores oficiais que serão utilizados "
        "nas conferências."
    )

    valor_qr = st.number_input(
        "Valor QR (R$)",
        min_value=0.0,
        value=7.00,
        step=0.01
    )

    valor_qs = st.number_input(
        "Valor QS (R$)",
        min_value=0.0,
        value=10.00,
        step=0.01
    )

    valor_diesel = st.number_input(
        "Diesel - R$/litro",
        min_value=0.0,
        value=0.0,
        step=0.01,
        help="Deixe 0 para apenas extrair o valor do documento."
    )

    valor_gasolina = st.number_input(
        "Gasolina - R$/litro",
        min_value=0.0,
        value=0.0,
        step=0.01,
        help="Deixe 0 para apenas extrair o valor do documento."
    )

    tolerancia = st.number_input(
        "Tolerância da conferência (R$)",
        min_value=0.0,
        value=0.05,
        step=0.01
    )


# ============================================================
# UPLOAD
# ============================================================

arquivo = st.file_uploader(
    "Selecione o P Trab",
    type=["pdf", "xlsx"]
)

if arquivo is None:
    st.stop()


# ============================================================
# LEITURA
# ============================================================

try:
    extensao = arquivo.name.lower().split(".")[-1]

    if extensao == "pdf":

        with st.spinner("Lendo o PDF..."):
            texto, paginas = ler_pdf(arquivo)

        st.success(
            f"📕 PDF identificado — {len(paginas)} página(s)."
        )

        if not texto.strip():
            st.error(
                "O PDF não possui texto pesquisável. "
                "Pode ser um PDF digitalizado."
            )
            st.stop()

    elif extensao == "xlsx":

        with st.spinner("Lendo a planilha..."):
            texto, abas = ler_excel(arquivo)

        st.success(
            f"📗 Excel identificado — {len(abas)} aba(s)."
        )

    else:
        st.error("Formato não suportado.")
        st.stop()

except Exception as erro:
    st.error("Não foi possível ler o arquivo.")
    st.exception(erro)
    st.stop()


# ============================================================
# ANÁLISE
# ============================================================

with st.spinner("Analisando o P Trab..."):
    resultado = analisar(texto)


# ============================================================
# RESUMO
# ============================================================

st.divider()

st.header("📋 Resultado da análise")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Fase",
        resultado["fase"]
    )

with col2:
    st.markdown(
        f"""
        <div style="text-align: center;">
            <p style="font-size: 14px; margin-bottom: 4px;">
                Classes
            </p>
            <p style="font-size: 32px; margin: 0;">
                {len(resultado["classes"])}
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

with col3:
    st.metric(
        "Efetivos encontrados",
        len(resultado["efetivos"])
    )

with col4:
    st.metric(
        "Valores encontrados",
        len(resultado["valores"])
    )


# ============================================================
# FASE
# ============================================================

st.subheader("🎯 Tipo de Plano")

if resultado["fase"] == "Não identificado":
    st.warning(
        "Não foi possível identificar automaticamente "
        "se o documento é Preparo ou Emprego."
    )
else:
    st.success(resultado["fase"])


# ============================================================
# CLASSES
# ============================================================

st.subheader("🏷️ Classes identificadas")

if resultado["classes"]:

    for classe in resultado["classes"]:

        if classe == "CLASSE I":
            st.write("🍽️ **CLASSE I — Alimentação**")

        elif classe == "CLASSE III":
            st.write("⛽ **CLASSE III — Combustíveis**")

        elif classe == "CLASSE VI":
            st.write(
                "🛠️ **CLASSE VI — Materiais/Serviços**"
            )

        elif classe == "CLASSE IX":
            st.write("🚙 **CLASSE IX — Manutenção**")

else:
    st.warning(
        "Nenhuma das classes configuradas foi identificada."
    )


# ============================================================
# DADOS GERAIS
# ============================================================

st.subheader("👥 Efetivo")

if resultado["efetivos"]:
    st.write(
        ", ".join(
            f"{x} militar(es)"
            for x in resultado["efetivos"]
        )
    )
else:
    st.write("Não identificado.")


st.subheader("📅 Quantidade de dias")

if resultado["dias"]:
    st.write(
        ", ".join(
            f"{x} dia(s)"
            for x in resultado["dias"]
        )
    )
else:
    st.write("Não identificado.")


st.subheader("🏦 CODUG / UGE")

if resultado["codug"]:
    st.write(", ".join(resultado["codug"]))
else:
    st.write("Não identificado.")


st.subheader("🧾 Natureza da Despesa")

if resultado["nd"]:
    st.write(", ".join(resultado["nd"]))
else:
    st.write("Não identificada.")


# ============================================================
# CLASSE I
# ============================================================

if "CLASSE I" in resultado["classes"]:

    st.divider()
    st.header("🍽️ Classe I — Alimentação")

    alimentacao = resultado["alimentacao"]

    if alimentacao:
        st.write(
            "**Modalidades encontradas:** "
            + ", ".join(alimentacao)
        )

    else:
        st.write(
            "Modalidade QR/QS não identificada."
        )

    st.write("**Valores de referência configurados:**")

    c1, c2 = st.columns(2)

    c1.metric(
        "QR",
        moeda(valor_qr)
    )

    c2.metric(
        "QS",
        moeda(valor_qs)
    )

    trecho = trecho_classe(
        texto,
        "CLASSE I"
    )

    with st.expander(
        "🔎 Ver trechos relacionados à Classe I"
    ):
        st.text(trecho or "Nenhum trecho localizado.")


# ============================================================
# CLASSE III
# ============================================================

if "CLASSE III" in resultado["classes"]:

    st.divider()
    st.header("⛽ Classe III — Combustíveis")

    combustiveis, litros = resultado["combustiveis"]

    if combustiveis:
        st.write(
            "**Combustível identificado:** "
            + ", ".join(combustiveis)
        )

    if litros:
        st.write("**Quantidades em litros encontradas:**")

        df_litros = pd.DataFrame({
            "Litros": litros
        })

        st.dataframe(
            df_litros,
            use_container_width=True,
            hide_index=True
        )

    if valor_diesel > 0 and litros:
        st.write("### Simulação com valor do Diesel")

        tabela = []

        for quantidade in litros:
            total = quantidade * valor_diesel

            tabela.append({
                "Litros": quantidade,
                "Valor unitário": valor_diesel,
                "Total calculado": total
            })

        st.dataframe(
            pd.DataFrame(tabela),
            use_container_width=True,
            hide_index=True
        )

    trecho = trecho_classe(
        texto,
        "CLASSE III"
    )

    with st.expander(
        "🔎 Ver trechos relacionados à Classe III"
    ):
        st.text(trecho or "Nenhum trecho localizado.")


# ============================================================
# CLASSE VI
# ============================================================

if "CLASSE VI" in resultado["classes"]:

    st.divider()
    st.header("🛠️ Classe VI — Material/Serviços")

    trecho = trecho_classe(
        texto,
        "CLASSE VI"
    )

    with st.expander(
        "🔎 Ver dados da Classe VI"
    ):
        st.text(trecho or "Nenhum trecho localizado.")


# ============================================================
# CLASSE IX
# ============================================================

if "CLASSE IX" in resultado["classes"]:

    st.divider()
    st.header("🚙 Classe IX — Manutenção")

    trecho = trecho_classe(
        texto,
        "CLASSE IX"
    )

    with st.expander(
        "🔎 Ver dados da Classe IX"
    ):
        st.text(trecho or "Nenhum trecho localizado.")


# ============================================================
# VALORES FINANCEIROS
# ============================================================

st.divider()
st.header("💰 Valores encontrados no documento")

if resultado["valores"]:

    df_valores = pd.DataFrame({
        "Valor encontrado": resultado["valores"]
    })

    df_valores["Valor formatado"] = (
        df_valores["Valor encontrado"].apply(moeda)
    )

    st.dataframe(
        df_valores[["Valor formatado"]],
        use_container_width=True,
        hide_index=True
    )

else:
    st.warning(
        "Nenhum valor no formato R$ foi localizado."
    )


# ============================================================
# RESUMO POR CLASSE
# ============================================================

st.divider()
st.header("📊 Resumo das Classes")

linhas_resumo = []

for classe in resultado["classes"]:

    if classe == "CLASSE I":
        descricao = "Alimentação"

    elif classe == "CLASSE III":
        descricao = "Combustíveis"

    elif classe == "CLASSE VI":
        descricao = "Material/Serviços"

    elif classe == "CLASSE IX":
        descricao = "Manutenção"

    else:
        descricao = ""

    linhas_resumo.append({
        "Classe": classe,
        "Descrição": descricao,
        "Situação": "Identificada"
    })


if linhas_resumo:

    df_resumo = pd.DataFrame(linhas_resumo)

    st.dataframe(
        df_resumo,
        use_container_width=True,
        hide_index=True
    )

else:
    st.warning("Nenhuma classe identificada.")


# ============================================================
# CONFERÊNCIA AUTOMÁTICA SIMPLES DA CLASSE I
# ============================================================

st.divider()
st.header("🧮 Memória de cálculo")

st.caption(
    "Os cálculos abaixo somente são realizados quando o "
    "sistema encontra dados suficientes no documento."
)

calculos = []

if resultado["efetivos"] and resultado["dias"]:

    efetivo = resultado["efetivos"][0]
    dias = resultado["dias"][0]

    if "QR" in resultado["alimentacao"]:

        total_qr = efetivo * valor_qr * dias

        calculos.append({
            "Classe": "I",
            "Item": "QR",
            "Efetivo": efetivo,
            "Dias": dias,
            "Valor unitário": valor_qr,
            "Total calculado": total_qr,
            "Memória": (
                f"{efetivo} × R$ {valor_qr:.2f} × "
                f"{dias} = {moeda(total_qr)}"
            )
        })

    if "QS" in resultado["alimentacao"]:

        total_qs = efetivo * valor_qs * dias

        calculos.append({
            "Classe": "I",
            "Item": "QS",
            "Efetivo": efetivo,
            "Dias": dias,
            "Valor unitário": valor_qs,
            "Total calculado": total_qs,
            "Memória": (
                f"{efetivo} × R$ {valor_qs:.2f} × "
                f"{dias} = {moeda(total_qs)}"
            )
        })


if calculos:

    df_calculos = pd.DataFrame(calculos)

    st.dataframe(
        df_calculos,
        use_container_width=True,
        hide_index=True
    )

    total_calculado = df_calculos[
        "Total calculado"
    ].sum()

    st.metric(
        "Total calculado da Classe I",
        moeda(total_calculado)
    )

else:
    st.info(
        "Ainda não há dados suficientes para montar "
        "automaticamente a memória de cálculo."
    )


# ============================================================
# DOWNLOAD DO RELATÓRIO CSV
# ============================================================

st.divider()
st.header("📥 Relatório")

dados_relatorio = []

for classe in resultado["classes"]:

    dados_relatorio.append({
        "Arquivo": arquivo.name,
        "Fase": resultado["fase"],
        "Classe": classe,
        "Efetivos": ", ".join(
            map(str, resultado["efetivos"])
        ),
        "Dias": ", ".join(
            map(str, resultado["dias"])
        ),
        "CODUG": ", ".join(resultado["codug"]),
        "ND": ", ".join(resultado["nd"])
    })


if dados_relatorio:

    df_relatorio = pd.DataFrame(dados_relatorio)

    csv = df_relatorio.to_csv(
        index=False,
        sep=";",
        encoding="utf-8-sig"
    )

    st.download_button(
        "⬇️ Baixar resumo da análise",
        data=csv,
        file_name="QUIXABA_resultado.csv",
        mime="text/csv"
    )


# ============================================================
# EXIBIÇÃO DO EXCEL
# ============================================================

if extensao == "xlsx":

    st.divider()
    st.header("📑 Abas da planilha")

    nome_aba = st.selectbox(
        "Selecione uma aba:",
        list(abas.keys())
    )

    st.dataframe(
        abas[nome_aba],
        use_container_width=True
    )


# ============================================================
# TEXTO EXTRAÍDO
# ============================================================

st.divider()

with st.expander(
    "🔎 Visualizar conteúdo completo extraído"
):
    st.text(texto)


# ============================================================
# FINAL
# ============================================================

st.success(
    "✅ Análise concluída pelo QUIXABA."
)

# ============================================================
# CONFERÊNCIA AUTOMÁTICA QUIXABA
# ============================================================

st.divider()
st.header("🧮 Conferência automática")

if arquivo is not None:

    # Valores definidos na lateral do sistema
    valores_referencia = {
        "QR": valor_qr,
        "QS": valor_qs,
        "DIESEL": valor_diesel,
        "GASOLINA": valor_gasolina,
    }

    st.write("### 📋 Valores utilizados na conferência")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("QR", f"R$ {valor_qr:.2f}")
    col2.metric("QS", f"R$ {valor_qs:.2f}")

    if valor_diesel > 0:
        col3.metric("Diesel", f"R$ {valor_diesel:.2f}/L")
    else:
        col3.metric("Diesel", "Valor do documento")

    if valor_gasolina > 0:
        col4.metric("Gasolina", f"R$ {valor_gasolina:.2f}/L")
    else:
        col4.metric("Gasolina", "Valor do documento")

    st.info(
        "O QUIXABA utilizará esses valores para refazer os cálculos "
        "e comparar com os valores encontrados no P Trab."
    )