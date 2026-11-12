import streamlit as st
import pandas as pd
import io
import re

st.set_page_config(page_title="Conversor CI / PL", page_icon="📊", layout="wide")

st.title("📊 Tratamento de Planilha - CI & PL")
st.write("Anexe a planilha de origem contendo as abas **CI** e **PL** para gerar o arquivo tratado.")

def formatar_numero_br(valor):
    """Converte valores numéricos ou strings com ponto para formato string com vírgula (padrão BR)."""
    if pd.isna(valor) or valor == "":
        return ""
    try:
        val_str = str(valor).strip().replace("$", "").replace(" ", "")
        if "," in val_str and "." in val_str:
            val_str = val_str.replace(",", "")
        num = float(val_str)
        return str(num).replace(".", ",")
    except ValueError:
        return str(valor).replace(".", ",")

def extrair_invoice(df_raw):
    """Procura o texto 'INVOICE NO.:' na planilha bruta e extrai o número logo após."""
    for row_idx in range(len(df_raw)):
        for col_idx in range(len(df_raw.columns)):
            val = str(df_raw.iloc[row_idx, col_idx]).strip()
            if "INVOICE NO.:" in val.upper():
                match = re.search(r"INVOICE\s*NO\.:\s*([^\s]+)", val, re.IGNORECASE)
                if match:
                    return match.group(1).strip().strip('"').strip("'")
                
                for next_col in range(col_idx + 1, len(df_raw.columns)):
                    next_val = df_raw.iloc[row_idx, next_col]
                    if pd.notna(next_val) and str(next_val).strip() != "":
                        return str(next_val).strip().strip('"').strip("'")
    return "NÃO ENCONTRADO"

def carregar_aba_inteligente(excel_data, sheet_name):
    """Carrega a aba unificando cabeçalhos divididos em linhas múltiplas."""
    df_raw = pd.read_excel(excel_data, sheet_name=sheet_name, header=None)
    
    # Procura nas primeiras 12 linhas a linha que contém o identificador 'code' ou 'material'
    header_idx = None
    for idx in range(min(12, len(df_raw))):
        row_vals = " ".join([str(v).strip().lower() for v in df_raw.iloc[idx].values if pd.notna(v)])
        if "code" in row_vals or "material" in row_vals:
            header_idx = idx
            break
            
    if header_idx is None:
        header_idx = 0

    # Se o cabeçalho for duplo (ex: MATERIAL em cima e CODE embaixo), une as duas linhas
    if header_idx + 1 < len(df_raw):
        cols = []
        for col_i in range(df_raw.shape[1]):
            val1 = str(df_raw.iloc[header_idx, col_i]).strip()
            val2 = str(df_raw.iloc[header_idx + 1, col_i]).strip()
            
            parts = []
            if val1 and val1.lower() != 'nan':
                parts.append(val1)
            if val2 and val2.lower() != 'nan' and val2.lower() not in parts:
                parts.append(val2)
            cols.append(" ".join(parts).upper())
            
        df = df_raw.iloc[header_idx + 2:].copy()
        df.columns = cols
    else:
        df = df_raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c).strip().upper() for c in df_raw.iloc[header_idx].values]
        
    return df, df_raw

uploaded_file = st.file_uploader("Selecione a planilha de origem (.xlsx ou .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        if "CI" not in excel_data.sheet_names or "PL" not in excel_data.sheet_names:
            st.error("A planilha precisa conter exatamente as abas **CI** e **PL**.")
            st.stop()

        # Carrega as abas com detecção inteligente de cabeçalho múltiplo
        df_ci, df_ci_raw = carregar_aba_inteligente(excel_data, "CI")
        df_pl, df_pl_raw = carregar_aba_inteligente(excel_data, "PL")

        # Extrai a invoice da planilha bruta
        invoice_num = extrair_invoice(df_ci_raw)
        if invoice_num == "NÃO ENCONTRADO":
            invoice_num = extrair_invoice(df_pl_raw)

        # Identificação flexível das colunas por palavras-chave
        def encontrar_coluna(colunas, termos):
            for c in colunas:
                c_clean = str(c).upper()
                if all(termo in c_clean for termo in termos):
                    return c
            return None

        col_mat_ci = encontrar_coluna(df_ci.columns, ["CODE"]) or encontrar_coluna(df_ci.columns, ["MATERIAL"])
        col_cip_ci = encontrar_coluna(df_ci.columns, ["CIP", "AMOUNT"]) or encontrar_coluna(df_ci.columns, ["AMOUNT"])

        if not col_mat_ci:
            st.error(f"Não foi possível encontrar a coluna de Código/Material na aba **CI**. Colunas detectadas: {list(df_ci.columns)}")
            st.stop()
        if not col_cip_ci:
            st.error(f"Não foi possível encontrar a coluna de Preço/CIP Amount na aba **CI**. Colunas detectadas: {list(df_ci.columns)}")
            st.stop()

        # Filtra linhas válidas da CI
        df_ci_valid = df_ci[df_ci[col_mat_ci].notna()].copy()
        df_ci_valid[col_mat_ci] = df_ci_valid[col_mat_ci].astype(str).str.strip()
        df_ci_valid = df_ci_valid[~df_ci_valid[col_mat_ci].str.lower().isin(["total", "nan", ""])]

        # Colunas da PL (Material Code, Total Quantity, Total N.W)
        col_mat_pl = encontrar_coluna(df_pl.columns, ["CODE"]) or encontrar_coluna(df_pl.columns, ["MATERIAL"]) or (df_pl.columns[24] if len(df_pl.columns) > 24 else None)
        col_nw_pl = encontrar_coluna(df_pl.columns.astype(str), ["N.W"]) or encontrar_coluna(df_pl.columns.astype(str), ["WEIGHT"]) or (df_pl.columns[10] if len(df_pl.columns) > 10 else None)
        col_qty_pl = encontrar_coluna(df_pl.columns.astype(str), ["QUANTITY"]) or encontrar_coluna(df_pl.columns.astype(str), ["QTY"]) or (df_pl.columns[8] if len(df_pl.columns) > 8 else None)

        if not col_mat_pl or not col_nw_pl or not col_qty_pl:
            st.error(f"Erro ao identificar colunas na aba **PL**. Verifique os nomes. Colunas detectadas: {list(df_pl.columns)}")
            st.stop()

        df_pl_valid = df_pl[df_pl[col_mat_pl].notna()].copy()
        df_pl_valid[col_mat_pl] = df_pl_valid[col_mat_pl].astype(str).str.strip()

        # Converte para numérico e agrupa caso haja duplicadas
        df_pl_valid[col_nw_pl] = pd.to_numeric(df_pl_valid[col_nw_pl].astype(str).str.replace(",", "."), errors="coerce").fillna(0)
        df_pl_valid[col_qty_pl] = pd.to_numeric(df_pl_valid[col_qty_pl].astype(str).str.replace(",", "."), errors="coerce").fillna(0)

        pl_agrupado = df_pl_valid.groupby(col_mat_pl, as_index=False).agg({
            col_nw_pl: "sum",
            col_qty_pl: "sum"
        })

        mapa_peso = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_nw_pl]))
        mapa_qtd = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_qty_pl]))

        # Monta o DataFrame final
        linhas_resultado = []
        for _, row in df_ci_valid.iterrows():
            partnumber = row[col_mat_ci]
            peso_bruto = mapa_peso.get(partnumber, "")
            qtd_bruta = mapa_qtd.get(partnumber, "")
            preco_bruto = row[col_cip_ci]

            if isinstance(qtd_bruta, float) and qtd_bruta.is_integer():
                qtd_formatada = int(qtd_bruta)
            else:
                qtd_formatada = qtd_bruta

            linhas_resultado.append({
                "PARTNUMBER": partnumber,
                "MOEDA": "220",
                "PESOUNITARIO": formatar_numero_br(peso_bruto),
                "QUANTIDADE": qtd_formatada,
                "PRECOTOTAL": formatar_numero_br(preco_bruto),
                "INVOICE": invoice_num,
                "INCOTERMS": "CIP"
            })

        df_final = pd.DataFrame(linhas_resultado)

        st.success(f"Processamento concluído! **Invoice identificada:** `{invoice_num}` | **Itens processados:** `{len(df_final)}`")
        st.dataframe(df_final, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Tratado")
        output.seek(0)

        st.download_button(
            label="📥 Baixar Planilha Tratada (.xlsx)",
            data=output,
            file_name=f"Planilha_Tratada_{invoice_num}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {e}")
