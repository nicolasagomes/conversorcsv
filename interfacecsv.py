import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="Conversor XLSX → CSV", page_icon="🧾", layout="centered")
st.title("🧾 Conversor de XLSX para CSV")
st.write("Faça upload de uma planilha Excel e baixe o CSV convertido.")

uploaded_file = st.file_uploader("Selecione seu arquivo Excel", type=["xlsx", "xls"])

def detectar_engine(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".xlsx":
        return "openpyxl"   # requer openpyxl
    elif ext == ".xls":
        return "xlrd"       # requer xlrd
    else:
        raise ValueError(f"Extensão não suportada: {ext}")

if uploaded_file is not None:
    try:
        engine = detectar_engine(uploaded_file.name)

        # Carrega o arquivo Excel e lista as abas com o engine correto
        xl = pd.ExcelFile(uploaded_file, engine=engine)
        sheet = st.selectbox("Selecione a aba (sheet)", xl.sheet_names)

        st.subheader("Opções de exportação")
        col1, col2 = st.columns(2)
        with col1:
            sep = st.selectbox(
                "Delimitador do CSV",
                options=[",", ";", "|", "\\t"],
                index=1,
                format_func=lambda x: "\\t (TAB)" if x == "\\t" else x
            )
        with col2:
            encoding = st.selectbox(
                "Codificação (encoding)",
                options=["utf-8-sig", "utf-8", "latin-1"],
                index=0
            )

        unir_colunas = st.checkbox(
            "Unir **todas as colunas** em **uma única coluna** no CSV",
            value=True,
            help="Cada linha vira uma única célula, com os valores concatenados pelo delimitador escolhido."
        )

        # Cabeçalho só é configurável quando NÃO for coluna única
        if unir_colunas:
            incluir_header = False  # força sem título
        else:
            incluir_header = st.checkbox("Incluir cabeçalho (nomes de colunas) no CSV", value=True)

        incluir_index = st.checkbox("Incluir índice no CSV", value=False)

        # Lê a aba escolhida com o mesmo engine
        df = pd.read_excel(xl, sheet_name=sheet, engine=engine)

        st.subheader("Pré-visualização")
        st.caption("Mostrando as primeiras 50 linhas.")
        st.dataframe(df.head(50), use_container_width=True)

        # Preparação do DataFrame de saída
        sep_effective = "\t" if sep == "\\t" else sep
        if unir_colunas:
            # Converte tudo para string e remove quebras de linha (melhor com regex, evita quebrar o CSV)
            df_str = df.astype(str).replace({r"[\r\n]+": " "}, regex=True)
            # Junta cada linha usando o delimitador
            serie_conc = df_str.agg(sep_effective.join, axis=1)
            df_out = pd.DataFrame({"_": serie_conc})  # header será omitido
        else:
            df_out = df

        # Gera CSV; para evitar erro de encoding, usamos errors='replace'
        csv_text = df_out.to_csv(index=incluir_index, header=incluir_header, sep=sep_effective)
        csv_bytes = csv_text.encode(encoding, errors="replace")

        original_name = uploaded_file.name.rsplit(".", 1)[0]
        suffix = "_coluna_unica" if unir_colunas else ""
        csv_filename = f"{original_name}{suffix}.csv"

        st.download_button(
            label="⬇️ Baixar CSV",
            data=csv_bytes,
            file_name=csv_filename,
            mime="text/csv"
        )

        with st.expander("Ver CSV (texto)"):
            st.code(csv_text[:20000])

    except ImportError as e:
        st.error(
            "Dependência ausente no servidor. Para `.xlsx` é preciso `openpyxl`; para `.xls`, `xlrd`.\n"
            f"Detalhe: {e}"
        )
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
else:
    st.info("Carregue um arquivo Excel para começar.")
