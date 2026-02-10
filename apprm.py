import streamlit as st
import pandas as pd
import re
import io

# Configuração da Página
st.set_page_config(
    page_title="Consolidador BRB",
    page_icon="💰",
    layout="wide"
)

# ==========================================
# FUNÇÕES DE PROCESSAMENTO
# ==========================================

def separar_cnpj_nome(texto):
    """Separa CNPJ e Nome de uma string bagunçada."""
    if not isinstance(texto, str):
        return None, texto
    
    # Regex para capturar CNPJ padrão
    padrao_cnpj = r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})'
    match = re.search(padrao_cnpj, texto)
    
    if match:
        cnpj = match.group(1)
        nome = texto.replace(cnpj, '').strip()
        return cnpj, nome
    else:
        return None, texto

def processar_planilha(arquivo_uploaded):
    """Processa um único arquivo carregado."""
    nome_arquivo = arquivo_uploaded.name
    
    # 1. Definir Banco e Código
    if '422-6' in nome_arquivo:
        banco_label = '422-6'
        cod_contabil = '3313'
    elif '558-4' in nome_arquivo:
        banco_label = '558-4'
        cod_contabil = '3314'
    else:
        banco_label = 'VERIFICAR_NOME_ARQUIVO'
        cod_contabil = ''

    try:
        df_raw = pd.read_excel(arquivo_uploaded, header=None)
    except Exception as e:
        st.error(f"Erro ao ler {nome_arquivo}: {e}")
        return pd.DataFrame()

    # Validação básica de tamanho
    if len(df_raw) < 9:
        return pd.DataFrame()
        
    # Dados começam na linha 9 (índice 8)
    df_data = df_raw.iloc[8:].reset_index(drop=True)

    lista_processada = []
    registro_atual = None

    for index, row in df_data.iterrows():
        valor_baixa = str(row[1]) if pd.notna(row[1]) else ""
        
        # Se tem Data na coluna Baixa -> Nova Linha
        if valor_baixa.strip() != "":
            if registro_atual:
                lista_processada.append(registro_atual)

            hist1 = str(row[9]) if pd.notna(row[9]) else ""
            hist2 = str(row[10]) if pd.notna(row[10]) else ""
            historico_completo = (hist1 + " " + hist2).strip()

            fornecedor_inicial = str(row[13]) if pd.notna(row[13]) else ""
            val_debito = row[18] if pd.notna(row[18]) else 0

            registro_atual = {
                'Baixa': row[1],
                'Emissao': row[2],
                'Cheq_Doc': row[4],
                'Natureza': row[5],
                'Historico': historico_completo,
                'Centro_Responsabilidade': row[12],
                'Fornecedor_Raw': fornecedor_inicial,
                'Credito': 0,
                'Debito': val_debito,
                'BANCO': banco_label,
                'CODIGO_CONTABIL': cod_contabil,
                'Arquivo_Origem': nome_arquivo
            }
        
        else:
            # Continuação de Fornecedor quebrado na linha de baixo
            texto_extra = str(row[13]) if pd.notna(row[13]) else ""
            if registro_atual and texto_extra.strip() != "":
                registro_atual['Fornecedor_Raw'] += " " + texto_extra.strip()

    if registro_atual:
        lista_processada.append(registro_atual)

    return pd.DataFrame(lista_processada)

def converter_df_para_excel(df):
    """Gera o binário do Excel para download usando openpyxl."""
    output = io.BytesIO()
    # AQUI ESTAVA O ERRO: Mudei engine='xlsxwriter' para 'openpyxl'
    # 'openpyxl' já vem instalado com o pandas por padrão
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Consolidado')
    processed_data = output.getvalue()
    return processed_data

# ==========================================
# LAYOUT (FRONT-END)
# ==========================================

st.title("📑 Consolidador de Extratos BRB")

with st.expander("ℹ️ Como usar (Clique para ler)", expanded=True):
    st.markdown("""
    1. **Arraste** os arquivos do BRB para a área abaixo.
    2. O sistema vai ler, **corrigir** os fornecedores e **juntar** tudo.
    3. Verifique a tabela de pré-visualização.
    4. Clique no botão **"Baixar Planilha"** que aparecerá no final.
    
    *Atenção: Os arquivos precisam ter '422-6' ou '558-4' no nome.*
    """)

# Upload
uploaded_files = st.file_uploader(
    "Solte os arquivos Excel aqui:", 
    type=['xlsx', 'xls'], 
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("🚀 Processar Arquivos", type="primary"):
        with st.spinner('Processando... Aguarde um instante.'):
            dfs_consolidados = []
            
            # Barra de progresso
            bar = st.progress(0)
            
            for i, file in enumerate(uploaded_files):
                df_temp = processar_planilha(file)
                if not df_temp.empty:
                    dfs_consolidados.append(df_temp)
                bar.progress((i + 1) / len(uploaded_files))

            if dfs_consolidados:
                # Junta tudo
                df_final = pd.concat(dfs_consolidados, ignore_index=True)

                # Separa CNPJ/Nome
                df_final[['CNPJ', 'Nome']] = df_final['Fornecedor_Raw'].apply(
                    lambda x: pd.Series(separar_cnpj_nome(x))
                )

                # Organiza colunas
                colunas_ordem = [
                    'Baixa', 'Emissao', 'Cheq_Doc', 'Natureza', 'Historico', 
                    'Centro_Responsabilidade', 'CNPJ', 'Nome', 
                    'Credito', 'Debito', 'BANCO', 'CODIGO_CONTABIL', 'Arquivo_Origem'
                ]
                
                # Garante que não quebre se faltar coluna
                for col in colunas_ordem:
                    if col not in df_final.columns:
                        df_final[col] = None
                        
                df_final = df_final[colunas_ordem]

                bar.empty() # Remove a barra de progresso
                st.success(f"✅ Sucesso! {len(df_final)} linhas geradas.")
                
                # Exibe prévia
                st.subheader("Pré-visualização (Primeiras 50 linhas)")
                st.dataframe(df_final.head(50), use_container_width=True)

                # --- Botão de Download ---
                # A função agora usa 'openpyxl', corrigindo o erro
                excel_data = converter_df_para_excel(df_final)
                
                st.download_button(
                    label="📥 BAIXAR PLANILHA CONSOLIDADA",
                    data=excel_data,
                    file_name="Consolidado_BRB_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

            else:
                st.error("❌ Não foi possível ler dados válidos. Verifique se as planilhas estão corretas.")