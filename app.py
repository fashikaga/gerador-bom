import streamlit as st
import pandas as pd

# Configuração da página da aplicação do Cliente
st.set_page_config(page_title="Configurador de Rádios TETRA | Motorola Solutions", layout="wide", page_icon="📻")

# CSS personalizado para layout profissional
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    .card-radio {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        border-left: 5px solid #0056b3;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 20px;
    }
    .badge-tag {
        background-color: #e9ecef;
        color: #495057;
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 0.85rem;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# 1. Carregar planilha LPU ABIX V4
@st.cache_data
def load_lpu():
    df = pd.read_excel("LPU ABIX V4.xlsx", sheet_name="LPU ABIX")
    df = df.dropna(subset=['PN', 'Categoria'])
    return df

df_lpu = load_lpu()

# 2. Mapeamento de Fichas e Resumos dos Rádios
INFO_RADIOS = {
    'Portátil - MTP3500': {
        'nome': 'MTP3500 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MTP3000',
        'resumo': 'Rádio portátil TETRA robusto, display colorido parcial, teclado limitado. Ideal para operação essencial em campo com Bluetooth e GPS.',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Simplificado (Limited Keypad)',
        'protecao': 'IP65 / IP67'
    },
    'Portátil - MTP3550': {
        'nome': 'MTP3550 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MTP3000',
        'resumo': 'Rádio portátil TETRA completo com teclado alfanumérico e display colorido. Desenvolvido para usuários que demandam chamadas diretas e mensagens rápidas.',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Alfanumérico Completo (Full Keypad)',
        'protecao': 'IP65 / IP67'
    },
    'Portátil - MXP600': {
        'nome': 'MXP600 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MXP600/MXP660',
        'resumo': 'Portátil TETRA de última geração. Leve, compacto, áudio de altíssima fidelidade com supressão de ruído por IA, Wi-Fi e Bluetooth 5.0.',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Completo com Display HD',
        'protecao': 'IP68 (Submersível)'
    },
    'Portátil - MXP660': {
        'nome': 'MXP660 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MXP600/MXP660',
        'resumo': 'Versão avançada do MXP600 com LTE/Wi-Fi ready e funções de missão crítica aprimoradas para ambientes de alta complexidade.',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Completo',
        'protecao': 'IP68'
    },
    'Portátil - MTP8550EX - 380MHz': {
        'nome': 'MTP8550EX (ATEX / Intrinsecamente Seguro)',
        'tipo': 'Portátil ATEX',
        'compatibilidade_acessorio': 'MTP8550EX',
        'resumo': 'Rádio portátil intrinsecamente seguro (ATEX/IECEx) projetado para ambientes explosivos (Óleo & Gás, Química). Inclui Man Down e GPS.',
        'frequencia': '380 - 430 MHz / 800 MHz',
        'teclado': 'Teclado Completo (Apto para uso com luvas)',
        'protecao': 'IP66 / IP67 / ATEX Zone 1 & 21'
    },
    'Móvel - MXM600': {
        'nome': 'MXM600 TETRA (Móvel / Fixo)',
        'tipo': 'Móvel / Veicular',
        'compatibilidade_acessorio': 'MXM600',
        'resumo': 'Rádio móvel e veicular TETRA de alta potência. Conectividade flexível em veículos, estações de despacho e maletas de emergência.',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Cabeçote de Controle Digital',
        'protecao': 'IP54 / Mil-Std 810G'
    }
}

# --- HEADER DA PÁGINA ---
st.title("📻 Configurador de Soluções de Rádio TETRA")
st.caption("Selecione os terminais portáteis, móveis e acessórios adequados para a sua operação.")

# Sidebar de navegação/resumo
st.sidebar.header("🛒 Sua Seleção")
carrinho_radios = []
carrinho_acessorios = []

# --- 1. SELEÇÃO DE RÁDIOS ---
st.header("1. Escolha dos Terminais (Rádios Portáteis e Móveis)")

for cat_key, info in INFO_RADIOS.items():
    df_radio = df_lpu[df_lpu['Categoria'] == cat_key]
    
    if not df_radio.empty:
        with st.container():
            st.markdown(f"""
            <div class="card-radio">
                <h3>{info['nome']} <span class="badge-tag">{info['tipo']}</span></h3>
                <p>{info['resumo']}</p>
                <small><b>Frequência:</b> {info['frequencia']} | <b>Teclado:</b> {info['teclado']} | <b>Proteção:</b> {info['protecao']}</small>
            </div>
            """, unsafe_allow_html=True)
            
            col_check, col_qtd = st.columns([1, 2])
            with col_check:
                selecionado = st.checkbox(f"Selecionar {info['nome']}", key=f"sel_{cat_key}")
            
            if selecionado:
                with col_qtd:
                    qtd = st.number_input(f"Quantidade de {info['nome']}", min_value=1, value=1, key=f"qtd_{cat_key}")
                    carrinho_radios.append({
                        'categoria': cat_key,
                        'nome': info['nome'],
                        'qtd': qtd,
                        'compat_acessorio': info['compatibilidade_acessorio']
                    })
            st.divider()

# --- 2. SELEÇÃO DE ACESSÓRIOS COMPATÍVEIS ---
st.header("2. Acessórios Compatíveis")

if carrinho_radios:
    # Identifica quais famílias de acessórios ativar com base nos rádios escolhidos
    familias_compativeis = set([r['compat_acessorio'] for r in carrinho_radios])
    
    # Filtra planilha de 'Portátil Acessórios'
    df_acessorios = df_lpu[df_lpu['Categoria'] == 'Portátil Acessórios'].copy()
    
    # Mapear e filtrar apenas acessórios das famílias ativas
    acessorios_disponiveis = []
    
    for _, row in df_acessorios.iterrows():
        grupo_str = str(row['Grupo'])
        for fam in familias_compativeis:
            if fam in grupo_str:
                acessorios_disponiveis.append(row)
                break
                
    if acessorios_disponiveis:
        df_acess_filtrados = pd.DataFrame(acessorios_disponiveis)
        
        st.info("Abaixo estão listados apenas os acessórios compatíveis com os rádios selecionados acima.")
        
        # Agrupa por item de acessório
        for idx, acc in df_acess_filtrados.iterrows():
            c1, c2 = st.columns([3, 1])
            with c1:
                st.write(f"• **{acc['Descritivo']}** *(Compatibilidade: {acc['Grupo']})*")
            with c2:
                qtd_acc = st.number_input(f"Qtd", min_value=0, value=0, key=f"acc_{acc['PN']}_{idx}")
                if qtd_acc > 0:
                    carrinho_acessorios.append({
                        'PN': acc['PN'],
                        'Descritivo': acc['Descritivo'],
                        'Qtd': qtd_acc,
                        'Grupo': acc['Grupo']
                    })
    else:
        st.warning("Nenhum acessório encontrado para a combinação selecionada.")
else:
    st.warning("Selecione ao menos um rádio no passo 1 para visualizar os acessórios compatíveis.")

# --- 3. RESUMO DA SELEÇÃO PARA O CLIENTE ---
st.header("3. Resumo da sua Configuração")

if carrinho_radios:
    st.subheader("Rádios Selecionados")
    df_resumo_radios = pd.DataFrame([{'Modelo': r['nome'], 'Quantidade (Unid.)': r['qtd']} for r in carrinho_radios])
    st.table(df_resumo_radios)
    
    if carrinho_acessorios:
        st.subheader("Acessórios Adicionados")
        df_resumo_acc = pd.DataFrame([{'Acessório / Item': a['Descritivo'], 'Quantidade (Unid.)': a['Qtd']} for a in carrinho_acessorios])
        st.table(df_resumo_acc)
    
    # Botão para solicitar proposta comercial
    st.success("Configuração concluída! Clique no botão abaixo para enviar o resumo para o setor comercial.")
    
    with st.form("form_contato"):
        st.write("### Solicitado por:")
        c_nome, c_email, c_emp = st.columns(3)
        c_nome.text_input("Nome Completo")
        c_email.text_input("E-mail Corporativo")
        c_emp.text_input("Empresa")
        
        if st.form_submit_button("📩 Enviar Solicitação de Cotação"):
            st.balloons()
            st.success("Sua solicitação foi enviada com sucesso! Em breve um especialista da ABIX / Motorola entrará em contato.")