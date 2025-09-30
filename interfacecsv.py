import streamlit as st
import pandas as pd
import os
import csv
from io import BytesIO

# Configuração da página
st.set_page_config(page_title="Conversores de Arquivos", page_icon="🔄", layout="centered")

# Menu lateral
st.sidebar.image('logoarea.png')
st.sidebar.title("Menu")
page = st.sidebar.radio("Escolha uma funcionalidade:", ["XLSX para CSV", "CSV para XLSX"])


st.sidebar.markdown("<div style='height: 150px;'></div>", unsafe_allow_html=True)

# Footer fixo
st.sidebar.markdown("""
    <hr style='margin-top: 20px; margin-bottom: 5px;'>
    <div style='text-align: center; font-size: 12px; color: gray;'>
        ⚙️ Desenvolvido por <b>Nicolas Gomes</b>
    </div>
""", unsafe_allow_html=True)


# Função para detectar engine do Excel
def detectar_engine(filename: str) -> str:
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".xlsx":
        return "openpyxl"
    elif ext == ".xls":
        return "xlrd"
    else:
        raise ValueError(f"Extensão não suportada: {ext}")

# Função para detectar delimitador de forma robusta
def detectar_delimitador(csv_file):
    sample = csv_file.read(2048).decode("utf-8", errors="replace")
    csv_file.seek(0)
    sniffer = csv.Sniffer()
    try:
        dialect = sniffer.sniff(sample)
        return dialect.delimiter
    except csv.Error:
        return ","  # fallback padrão

# Função para ler CSV corretamente com cabeçalho
def ler_csv_com_cabecalho(csv_file):
    delimiter = detectar_delimitador(csv_file)
    df = pd.read_csv(csv_file, delimiter=delimiter, dtype=str, keep_default_na=False, na_values=[""])
    df.fillna("", inplace=True)
    return df

# Página 1: XLSX → CSV
if page == "XLSX para CSV":
    st.title("🧾 Conversor de XLSX para CSV")
    st.write("Faça upload de uma planilha Excel e baixe o CSV convertido.")

    uploaded_file = st.file_uploader("Selecione seu arquivo Excel", type=["xlsx", "xls"])

    if uploaded_file is not None:
        try:
            engine = detectar_engine(uploaded_file.name)
            xl = pd.ExcelFile(uploaded_file, engine=engine)
            sheet = st.selectbox("Selecione a aba (sheet)", xl.sheet_names)

            st.subheader("Opções de exportação")
            col1, col2 = st.columns(2)
            with col1:
                sep = st.selectbox("Delimitador do CSV", [",", ";", "|", "\\t"], index=1,
                                   format_func=lambda x: "\\t (TAB)" if x == "\\t" else x)
            with col2:
                encoding = st.selectbox("Codificação (encoding)", ["utf-8-sig", "utf-8", "latin-1"], index=0)

            unir_colunas = st.checkbox("Unir todas as colunas em uma única coluna", value=True)
            incluir_index = st.checkbox("Incluir índice no CSV", value=False)
            incluir_header = not unir_colunas and st.checkbox("Incluir cabeçalho no CSV", value=True)

            df = pd.read_excel(xl, sheet_name=sheet, engine=engine)
            st.subheader("Pré-visualização")
            st.dataframe(df.head(50), use_container_width=True)

            sep_effective = "\t" if sep == "\\t" else sep

            if unir_colunas:
                df_str = df.astype(str).replace({r"[\r\n]+": " "}, regex=True)
                header_line = sep_effective.join(df.columns)
                data_lines = df_str.agg(sep_effective.join, axis=1)
                df_final = pd.DataFrame([header_line] + list(data_lines), columns=["Dados"])
                csv_text = df_final.to_csv(index=False, header=False, sep=sep_effective)
                csv_bytes = csv_text.encode(encoding, errors="replace")
                suffix = "_coluna_unica"
            else:
                csv_text = df.to_csv(index=incluir_index, header=incluir_header, sep=sep_effective)
                csv_bytes = csv_text.encode(encoding, errors="replace")
                suffix = ""

            original_name = uploaded_file.name.rsplit(".", 1)[0]
            csv_filename = f"{original_name}{suffix}.csv"

            st.download_button(
                label="⬇️ Baixar CSV",
                data=csv_bytes,
                file_name=csv_filename,
                mime="text/csv"
            )

            with st.expander("Ver CSV (texto)"):
                st.code(csv_text[:20000])

        except Exception as e:
            st.error(f"Erro ao processar o arquivo: {e}")
    else:
        st.info("Carregue um arquivo Excel para começar.")

# Página 2: CSV → XLSX Ajustado
elif page == "CSV para XLSX":
    st.title("📄 Conversor de CSV para XLSX Ajustado")
    st.write("Faça upload de um arquivo CSV e baixe o Excel formatado automaticamente.")

    uploaded_csv = st.file_uploader("Selecione seu arquivo CSV", type=["csv"])

    if uploaded_csv is not None:
        try:
            df = ler_csv_com_cabecalho(uploaded_csv)
            st.subheader("Pré-visualização do Excel ajustado")
            st.dataframe(df.head(50), use_container_width=True)

            output = BytesIO()
            df.to_excel(output, index=False, engine='openpyxl')
            excel_bytes = output.getvalue()

            excel_filename = uploaded_csv.name.rsplit(".", 1)[0] + "_ajustado.xlsx"

            st.download_button(
                "⬇️ Baixar Excel Ajustado",
                data=excel_bytes,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro ao processar o CSV: {e}")
    else:
        st.info("Carregue um arquivo CSV para começar.")