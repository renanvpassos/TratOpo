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
        # Substitui o ponto decimal por vírgula mantendo todas as casas decimais originais
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
    """Carrega a aba identificando a linha real do cabeçalho por pontuação de palavras-chave."""
    df_raw = pd.read_excel(excel_data, sheet_name=sheet_name, header=None)
    
    best_row = 0
    max_score = -1
    palavras_chave = ["code", "material", "amount", "price", "qty", "quantity", "weight", "n.w", "cip", "description"]
    
    for idx in range(min(30, len(df_raw))):
        row_vals = [str(v).strip().lower() for v in df_raw.iloc[idx].values if pd.notna(v)]
        score = sum(1 for p in palavras_chave if any(p in v for v in row_vals))
        tem_material_ou_code = any("code" in v or "material" in v for v in row_vals)
        
        if tem_material_ou_code and score > max_score:
            max_score = score
            best_row = idx
            
    header_idx = best_row if max_score > 0 else 0

    if header_idx + 1 < len(df_raw):
        next_row_vals = [str(v).strip().lower() for v in df_raw.iloc[header_idx + 1].values if pd.notna(v)]
        score_next = sum(1 for p in palavras_chave if any(p in v for v in next_row_vals))
        
        if score_next >= 1 or any(v in ["code", "usd", "amount", "quantity", "n.w"] for v in next_row_vals):
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
            return df, df_raw

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

        df_ci, df_ci_raw = carregar_aba_inteligente(excel_data, "CI")
        df_pl, df_pl_raw = carregar_aba_inteligente(excel_data, "PL")

        invoice_num = extrair_invoice(df_ci_raw)
        if invoice_num == "NÃO ENCONTRADO":
            invoice_num = extrair_invoice(df_pl_raw)

        def encontrar_coluna(colunas, termos_obrigatorios, termos_proibidos=None):
            """Busca flexível de colunas garantindo termos obrigatórios e rejeitando proibidos (ex: evitar 'Total N.W' se quiser apenas 'N.W')."""
            termos_proibidos = termos_proibidos or []
            for c in colunas:
                c_clean = str(c).upper()
                tem_obrigatorio = all(t in c_clean for t in termos_obrigatorios)
                tem_proibido = any(t in c_clean for t in termos_proibidos)
                if tem_obrigatorio and not tem_proibido:
                    return c
            return None

        # 1. Configurações da aba CI (PartNumber e Preço Total)
        col_mat_ci = encontrar_coluna(df_ci.columns, ["CODE"]) or encontrar_coluna(df_ci.columns, ["MATERIAL"])
        col_cip_ci = (
            encontrar_coluna(df_ci.columns, ["CIP", "AMOUNT"]) or 
            encontrar_coluna(df_ci.columns, ["AMOUNT"]) or 
            encontrar_coluna(df_ci.columns, ["PRICE"]) or
            encontrar_coluna(df_ci.columns, ["USD"])
        )

        if not col_mat_ci or not col_cip_ci:
            st.error(f"Erro ao identificar colunas na aba **CI**. Colunas detectadas: {list(df_ci.columns)}")
            st.stop()

        df_ci_valid = df_ci[df_ci[col_mat_ci].notna()].copy()
        df_ci_valid[col_mat_ci] = df_ci_valid[col_mat_ci].astype(str).str.strip()
        df_ci_valid = df_ci_valid[~df_ci_valid[col_mat_ci].str.lower().isin(["total", "nan", ""])]

        # 2. Configurações da aba PL (Material Code, N.W exato e Quantidade)
        col_mat_pl = encontrar_coluna(df_pl.columns, ["CODE"]) or encontrar_coluna(df_pl.columns, ["MATERIAL"]) or (df_pl.columns[24] if len(df_pl.columns) > 24 else None)
        
        # Procura por "N.W" estrito (evitando "TOTAL N.W" caso haja separação)
        col_nw_pl = (
            encontrar_coluna(df_pl.columns.astype(str), ["N.W"], termos_proibidos=["TOTAL"]) or 
            encontrar_coluna(df_pl.columns.astype(str), ["WEIGHT"], termos_proibidos=["TOTAL"]) or 
            (df_pl.columns[10] if len(df_pl.columns) > 10 else None)
        )
        
        col_qty_pl = (
            encontrar_coluna(df_pl.columns.astype(str), ["QUANTITY"]) or 
            encontrar_coluna(df_pl.columns.astype(str), ["QTY"]) or 
            (df_pl.columns[8] if len(df_pl.columns) > 8 else None)
        )

        if not col_mat_pl or not col_nw_pl or not col_qty_pl:
            st.error(f"Erro ao identificar colunas na aba **PL**. Colunas detectadas: {list(df_pl.columns)}")
            st.stop()

        df_pl_valid = df_pl[df_pl[col_mat_pl].notna()].copy()
        df_pl_valid[col_mat_pl] = df_pl_valid[col_mat_pl].astype(str).str.strip()

        # Limpeza e conversão numérica rigorosa para padrão americano (ponto decimal) para realizar a soma correta
        df_pl_valid[col_nw_pl] = pd.to_numeric(
            df_pl_valid[col_nw_pl].astype(str).str.replace(",", ".").str.replace(r"[^\d.-]", "", regex=True), 
            errors="coerce"
        ).fillna(0)
        
        df_pl_valid[col_qty_pl] = pd.to_numeric(
            df_pl_valid[col_qty_pl].astype(str).str.replace(",", ".").str.replace(r"[^\d.-]", "", regex=True), 
            errors="coerce"
        ).fillna(0)

        # Agrupa na PL somando os pesos e quantidades para cada PartNumber duplicado
        pl_agrupado = df_pl_valid.groupby(col_mat_pl, as_index=False).agg({
            col_nw_pl: "sum",
            col_qty_pl: "sum"
        })

        mapa_peso = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_nw_pl]))
        mapa_qtd = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_qty_pl]))

        # Monta o DataFrame final baseando-se nos itens da CI, puxando Peso Unitário e Quantidade somados da PL
        linhas_resultado = []
        for _, row in df_ci_valid.iterrows():
            partnumber = row[col_mat_ci]
            peso_bruto = mapa_peso.get(partnumber, 0)
            qtd_bruta = mapa_qtd.get(partnumber, 0)
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
