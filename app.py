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
        # Limpa possíveis espaços ou símbolos de moeda caso venha como texto
        val_str = str(valor).strip().replace("$", "").replace(" ", "")
        # Se já tiver vírgula e ponto (ex: 1,234.56), remove a vírgula de milhar
        if "," in val_str and "." in val_str:
            val_str = val_str.replace(",", "")
        num = float(val_str)
        # Formata com vírgula decimal (ajustável para quantas casas decimais precisar)
        return str(num).replace(".", ",")
    except ValueError:
        return str(valor).replace(".", ",")

def extrair_invoice(df_ci):
    """Procura o texto 'INVOICE NO.:' na aba CI e extrai o número logo após ou na célula vizinha."""
    for row_idx in range(len(df_ci)):
        for col_idx in range(len(df_ci.columns)):
            val = str(df_ci.iloc[row_idx, col_idx]).strip()
            if "INVOICE NO.:" in val.upper():
                # Caso 1: Está na mesma célula (ex: "INVOICE NO.: A12345678-000123")
                match = re.search(r"INVOICE\s*NO\.:\s*([^\s]+)", val, re.IGNORECASE)
                if match:
                    return match.group(1).strip().strip('"').strip("'")
                
                # Caso 2: O rótulo está em uma célula e o número na próxima coluna à direita
                for next_col in range(col_idx + 1, len(df_ci.columns)):
                    next_val = df_ci.iloc[row_idx, next_col]
                    if pd.notna(next_val) and str(next_val).strip() != "":
                        return str(next_val).strip().strip('"').strip("'")
    return "NÃO ENCONTRADO"

def encontrar_linha_cabecalho(df, nome_coluna_alvo):
    """Localiza em qual linha está o cabeçalho procurando pelo nome exato de uma coluna chave."""
    for idx, row in df.iterrows():
        valores = [str(v).strip().lower() for v in row.values if pd.notna(v)]
        if nome_coluna_alvo.lower() in valores:
            return idx
    return None

uploaded_file = st.file_uploader("Selecione a planilha de origem (.xlsx ou .xls)", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Lê as abas brutas primeiro para localizar cabeçalhos dinâmicos e a Invoice
        excel_data = pd.ExcelFile(uploaded_file)
        
        if "CI" not in excel_data.sheet_names or "PL" not in excel_data.sheet_names:
            st.error("A planilha precisa conter exatamente as abas **CI** e **PL**.")
            st.stop()

        df_ci_raw = pd.read_excel(excel_data, sheet_name="CI", header=None)
        df_pl_raw = pd.read_excel(excel_data, sheet_name="PL", header=None)

        # 1. Extrair INVOICE na aba CI (caso não ache, tenta na PL por garantia)
        invoice_num = extrair_invoice(df_ci_raw)
        if invoice_num == "NÃO ENCONTRADO":
            invoice_num = extrair_invoice(df_pl_raw)

        # 2. Localizar cabeçalhos das tabelas nas abas CI e PL
        header_ci = encontrar_linha_cabecalho(df_ci_raw, "Material Code")
        header_pl = encontrar_linha_cabecalho(df_pl_raw, "Material Code")

        if header_ci is None:
            st.error("Não foi possível encontrar a coluna 'Material Code' na aba **CI**.")
            st.stop()
        if header_pl is None:
            st.error("Não foi possível encontrar a coluna 'Material Code' na aba **PL**.")
            st.stop()

        # Recarrega as abas usando a linha correta como cabeçalho
        df_ci = pd.read_excel(excel_data, sheet_name="CI", header=header_ci)
        df_pl = pd.read_excel(excel_data, sheet_name="PL", header=header_pl)

        # Padroniza nomes das colunas (remove espaços extras nas pontas)
        df_ci.columns = [str(c).strip() for c in df_ci.columns]
        df_pl.columns = [str(c).strip() for c in df_pl.columns]

        # Valida colunas necessárias na aba CI
        col_mat_ci = next((c for c in df_ci.columns if c.lower() == "material code"), None)
        col_cip_ci = next((c for c in df_ci.columns if c.lower() == "cip amount usd"), None)

        if not col_mat_ci or not col_cip_ci:
            st.error(f"Colunas esperadas na aba CI não encontradas. Colunas atuais: {list(df_ci.columns)}")
            st.stop()

        # Filtra linhas válidas da CI (onde Material Code não é nulo e não é linha de Total)
        df_ci_valid = df_ci[df_ci[col_mat_ci].notna()].copy()
        df_ci_valid[col_mat_ci] = df_ci_valid[col_mat_ci].astype(str).str.strip()
        df_ci_valid = df_ci_valid[~df_ci_valid[col_mat_ci].str.lower().isin(["total", "nan", ""])]

        # Mapeamento da aba PL por índice de coluna fixa conforme especificação + Material Code:
        # Coluna I = índice 8 (Total Quantity) | Coluna K = índice 10 (Total N.W) | Coluna Y = índice 24 (Material Code)
        # Busca pelo nome primeiro; se falhar, usa o índice da letra da coluna (I=8, K=10, Y=24)
        col_mat_pl = next((c for c in df_pl.columns if c.lower() == "material code"), df_pl.columns[24] if len(df_pl.columns) > 24 else None)
        col_nw_pl = next((c for c in df_pl.columns if "total n.w" in c.lower()), df_pl.columns[10] if len(df_pl.columns) > 10 else None)
        col_qty_pl = next((c for c in df_pl.columns if "total quantity" in c.lower()), df_pl.columns[8] if len(df_pl.columns) > 8 else None)

        df_pl_valid = df_pl[df_pl[col_mat_pl].notna()].copy()
        df_pl_valid[col_mat_pl] = df_pl_valid[col_mat_pl].astype(str).str.strip()

        # Cria dicionários de consulta rápida (PartNumber -> Peso e Quantidade)
        # Caso haja mais de uma linha com o mesmo PartNumber na PL, soma os valores numéricos
        df_pl_valid[col_nw_pl] = pd.to_numeric(df_pl_valid[col_nw_pl], errors="coerce").fillna(0)
        df_pl_valid[col_qty_pl] = pd.to_numeric(df_pl_valid[col_qty_pl], errors="coerce").fillna(0)

        pl_agrupado = df_pl_valid.groupby(col_mat_pl, as_index=False).agg({
            col_nw_pl: "sum",
            col_qty_pl: "sum"
        })

        mapa_peso = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_nw_pl]))
        mapa_qtd = dict(zip(pl_agrupado[col_mat_pl], pl_agrupado[col_qty_pl]))

        # Monta o DataFrame final seguindo exatamente a ordem de A a G
        linhas_resultado = []
        for _, row in df_ci_valid.iterrows():
            partnumber = row[col_mat_ci]
            peso_bruto = mapa_peso.get(partnumber, "")
            qtd_bruta = mapa_qtd.get(partnumber, "")
            preco_bruto = row[col_cip_ci]

            # Formata quantidade como inteiro se não tiver casas decimais
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

        # Gera o arquivo Excel em memória para download
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
