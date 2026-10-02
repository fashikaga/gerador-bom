import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Gerador de BOM - ABIX", layout="wide")
st.title("📦 Configurador de Equipamentos e Gerador de BOM")

# 1. Carregar planilha LPU ABIX V4
@st.cache_data
def load_lpu():
    df = pd.read_excel("LPU ABIX V4.xlsx", sheet_name="LPU ABIX")
    # Remover linhas sem PN ou Categoria se houver
    df = df.dropna(subset=['PN', 'Categoria'])
    return df

df_lpu = load_lpu()

# 2. Seleção de Modelos pelo Cliente
st.sidebar.header("Seleção de Modelos")
categorias = df_lpu['Categoria'].unique()

itens_selecionados = []

st.subheader("1. Escolha os Modelos e Quantidades")

cols = st.columns(2)
col_idx = 0

for cat in categorias:
    df_cat = df_lpu[df_lpu['Categoria'] == cat]
    grupos = df_cat['Grupo'].unique()
    
    with cols[col_idx % 2]:
        st.markdown(f"### {cat}")
        for grupo in grupos:
            label = f"{cat} - Grupo {grupo}"
            # Checkbox para ativar o grupo/modelo
            ativo = st.checkbox(f"Incluir {label}", key=f"check_{cat}_{grupo}")
            if ativo:
                # Multiplicador (ex: quantas unidades do kit/grupo o cliente quer)
                qtd_multiplicador = st.number_input(
                    f"Quantidade de conjuntos para [{label}]", 
                    min_value=1, value=1, step=1, key=f"num_{cat}_{grupo}"
                )
                itens_selecionados.append({
                    'Categoria': cat,
                    'Grupo': grupo,
                    'Multiplicador': qtd_multiplicador
                })
        st.divider()
    col_idx += 1

# 3. Geração Automática da BOM
st.subheader("2. Lista BOM Gerada Automática")

if itens_selecionados:
    bom_rows = []
    
    for item in itens_selecionados:
        sub_df = df_lpu[(df_lpu['Categoria'] == item['Categoria']) & (df_lpu['Grupo'] == item['Grupo'])].copy()
        # Calcula a quantidade final do item na BOM
        sub_df['Qdade Total'] = sub_df['Qdade.'] * item['Multiplicador']
        bom_rows.append(sub_df)
    
    df_bom_bruta = pd.concat(bom_rows)
    
    # Agrupa por PN para consolidar duplicados
    df_bom_consolidada = df_bom_bruta.groupby(
        ['PN', 'Descritivo', 'Tipo', 'Importação', 'Unitário'], as_index=False
    ).agg({'Qdade Total': 'sum'})
    
    st.dataframe(df_bom_consolidada, use_container_width=True)
    
    # Exportar para Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_bom_consolidada.to_excel(writer, index=False, sheet_name='BOM Gerada')
    
    st.download_button(
        label="📥 Baixar Lista BOM (Excel)",
        data=buffer.getvalue(),
        file_name="BOM_Consolidada_Cliente.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("Selecione ao menos um modelo acima para gerar a lista BOM.")