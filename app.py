import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Configuração da página
st.set_page_config(page_title="Configurador de Soluções TETRA | Motorola Solutions", layout="wide", page_icon="📻")

# Estilização CSS
st.markdown("""
<style>
    .stApp { background-color: #f8f9fa; }
    .card-radio {
        background-color: #ffffff;
        padding: 22px;
        border-radius: 12px;
        border-left: 6px solid #0056b3;
        box-shadow: 0 4px 10px rgba(0,0,0,0.06);
        margin-bottom: 25px;
    }
    .badge-tag {
        background-color: #0056b3;
        color: #ffffff;
        padding: 4px 10px;
        border-radius: 4px;
        font-size: 0.82rem;
        font-weight: 600;
        float: right;
    }
    .feature-list {
        background-color: #f1f3f5;
        padding: 12px 15px;
        border-radius: 8px;
        margin-top: 10px;
        font-size: 0.9rem;
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

# 2. Informações Detalhadas dos Rádios Motorola Solutions
INFO_RADIOS = {
    'Portátil - MTP3500': {
        'nome': 'MTP3500 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MTP3000',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Simplificado (LKP - Limited Keypad)',
        'protecao': 'IP65 / IP67 (Resistente à água e poeira)',
        'resumo': 'Rádio portátil de missão crítica construído para alta durabilidade. Oferece clareza de áudio excepcional e operação intuitiva em ambientes severos.',
        'antena': 'Inclusa (Antena Whip 110mm / 380-430MHz)',
        'criptografia': 'TEA1 (Criptografia de Interface Aérea)',
        'bateria_padrão': 'Bateria Padrão Li-Ion 2200 mAh inclusa no kit base',
        'gps': 'Possui licença GPS ativada',
        'bluetooth': 'Bluetooth Smart Ready / Áudio sem fio',
        'potencia': 'Suporta Licença TOGGLE RF POWER CLASS 3 (Aumento de potência para 2.8W)',
        'mandown': 'Suporta Licença Man Down (Alarme/Sensor automático de queda do operador)'
    },
    'Portátil - MTP3550': {
        'nome': 'MTP3550 TETRA',
        'tipo': 'Portátil',
        'compatibilidade_acessorio': 'MTP3000',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Alfanumérico Completo (FKP - Full Keypad)',
        'protecao': 'IP65 / IP67',
        'resumo': 'Versão com teclado completo alfanumérico e display colorido. Permite envio e recebimento rápido de mensagens de texto, chamadas privadas e navegação completa por menus.',
        'antena': 'Inclusa (Antena Whip 110mm / 380-430MHz)',
        'criptografia': 'TEA1 (Criptografia de Interface Aérea)',
        'bateria_padrão': 'Bateria Padrão Li-Ion 2200 mAh inclusa',
        'gps': 'GPS Integrado com Licença Ativa',
        'bluetooth': 'Bluetooth Smart Ready & Localização Indoor',
        'potencia': 'Suporta Licença TOGGLE RF POWER CLASS 3 (Aumento de potência para 2.8W)',
        'mandown': 'Suporta Licença Man Down (Sensor de homem caído e imobilidade)'
    },
    'Portátil - MXP600': {
        'nome': 'MXP600 TETRA',
        'tipo': 'Portátil Avançado',
        'compatibilidade_acessorio': 'MXP600/MXP660',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Teclado Completo HD',
        'protecao': 'IP68 (Submersível até 2m por 2h)',
        'resumo': 'O terminal portátil TETRA mais leve e moderno do mercado. Possui supressão de ruído com Inteligência Artificial, Wi-Fi 2.4/5GHz e HD Audio.',
        'antena': 'Inclusa (Antena Whip UHF 380-470MHz)',
        'criptografia': 'TEA1 e End-to-End Encryption (E2EE) ready',
        'bateria_padrão': 'Bateria IMPRES 2 Li-Ion IP68 alta capacidade',
        'gps': 'Multi-GNSS (GPS, GLONASS, BEIDOU, GALILEO)',
        'bluetooth': 'Bluetooth 5.0 com localização indoor acelerada',
        'potencia': 'RF Class 3 (2.8W) estendido',
        'mandown': 'Man Down integrado com sensores de acelerômetro de última geração'
    },
    'Portátil - MTP8550EX - 380MHz': {
        'nome': 'MTP8550EX TETRA ATEX',
        'tipo': 'Portátil Intrinsecamente Seguro',
        'compatibilidade_acessorio': 'MTP8550EX',
        'frequencia': '380 - 430 MHz',
        'teclado': 'Teclado Completo (Projetado para uso com luvas pesadas)',
        'protecao': 'IP66 / IP67 & Certificação ATEX / IECEx Zona 1 e 21',
        'resumo': 'Projetado para atmosferas explosivas em indústrias químicas, refinarias e plataformas de petróleo. Formato em "T" ergonômico e display superior secundário.',
        'antena': 'Inclusa (Antena Whip ATEX 350-470MHz)',
        'criptografia': 'TEA1 Hardware Encryption',
        'bateria_padrão': 'Bateria IMPRES Li-Ion ATEX IP67 dedicada inclusa',
        'gps': 'GNSS Integrado',
        'bluetooth': 'Bluetooth ATEX seguro',
        'potencia': 'Potência otimizada para segurança intrínseca ATEX',
        'mandown': 'Sensor Man Down avançado de fábrica'
    },
    'Móvel - MXM600': {
        'nome': 'MXM600 TETRA',
        'tipo': 'Móvel / Veicular / Fixo',
        'compatibilidade_acessorio': 'MXM600',
        'frequencia': '350 - 470 MHz',
        'teclado': 'Cabeçote de Controle Inteligente',
        'protecao': 'IP54 / Mil-Std 810G',
        'resumo': 'Terminal veicular/fixo TETRA de alta potência operacional. Ideal para instalação em viaturas, ambulâncias, embarcações e salas de controle.',
        'antena': 'Conexão para Antena Veicular externa de alto ganho',
        'criptografia': 'TEA1 Criptografia de canal',
        'bateria_padrão': 'Alimentação via bateria do veículo / Fonte externa 12V',
        'gps': 'GPS/GNSS com suporte a antenas veiculares',
        'bluetooth': 'Bluetooth 5.0 integrado',
        'potencia': 'Potência nativa de rádio móvel veicular',
        'mandown': 'Não aplicável a rádios móveis (Emergência via botão/pedal externo)'
    }
}

# --- Função de Envio de E-mail ---
def enviar_email_cotacao(destinatario, dados_cliente, resumo_radios, resumo_sw, resumo_acc):
    try:
        # Obter credenciais salvas no Streamlit Secrets
        smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
        sender_email = st.secrets.get("SENDER_EMAIL", "")
        sender_password = st.secrets.get("SENDER_PASSWORD", "")

        if not sender_email or not sender_password:
            return False, "Configurações de e-mail (SENDER_EMAIL / SENDER_PASSWORD) não foram definidas no Streamlit Secrets."

        msg = MIMEMultipart()
        msg['From'] = f"Configurador TETRA <{sender_email}>"
        msg['To'] = destinatario
        msg['Subject'] = f"🚨 Nova Solicitação de Cotação TETRA - {dados_cliente['empresa']}"

        # Montagem do Corpo do E-mail em HTML
        html_content = f"""
        <h2>Nova Solicitação de Cotação de Equipamentos TETRA</h2>
        <hr>
        <h3>👤 Dados do Cliente / Solicitante:</h3>
        <ul>
            <li><b>Nome:</b> {dados_cliente['nome']}</li>
            <li><b>E-mail:</b> {dados_cliente['email']}</li>
            <li><b>Empresa:</b> {dados_cliente['empresa']}</li>
            <li><b>Telefone:</b> {dados_cliente['telefone']}</li>
        </ul>
        <hr>
        <h3>📻 Terminais / Rádios Selecionados:</h3>
        {pd.DataFrame(resumo_radios).to_html(index=False, border=1) if resumo_radios else '<p>Nenhum</p>'}
        <br>
        <h3>💻 Licenças de Software (SW Features):</h3>
        {pd.DataFrame(resumo_sw).to_html(index=False, border=1) if resumo_sw else '<p>Nenhuma licença extra selecionada</p>'}
        <br>
        <h3>🎧 Acessórios Adicionais:</h3>
        {pd.DataFrame(resumo_acc).to_html(index=False, border=1) if resumo_acc else '<p>Nenhum acessório extra selecionado</p>'}
        <hr>
        <p><small>Mensagem gerada automaticamente pelo Configurador Web TETRA ABIX / Motorola Solutions.</small></p>
        """

        msg.attach(MIMEText(html_content, 'html'))

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, destinatario, msg.as_string())
        server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e:
        return False, str(e)


# --- HEADER DA PÁGINA ---
st.title("📻 Configurador de Soluções de Rádio TETRA")
st.caption("Monte a sua solução de comunicação de missão crítica selecionando rádios, acessórios e licenças de software.")

carrinho_radios = []
carrinho_acessorios = []
carrinho_sw = []

# --- 1. SELEÇÃO DE RÁDIOS ---
st.header("1. Seleção dos Terminais (Rádios Portáteis e Móveis)")

for cat_key, info in INFO_RADIOS.items():
    df_radio = df_lpu[df_lpu['Categoria'] == cat_key]
    
    if not df_radio.empty:
        with st.container():
            st.markdown(f"""
            <div class="card-radio">
                <span class="badge-tag">{info['tipo']}</span>
                <h3>{info['nome']}</h3>
                <p>{info['resumo']}</p>
                <div class="feature-list">
                    <b>Frequência:</b> {info['frequencia']} | <b>Teclado:</b> {info['teclado']}<br>
                    <b>📡 Antena:</b> {info['antena']}<br>
                    <b>🔐 Criptografia:</b> {info['criptografia']}<br>
                    <b>🔋 Bateria do Kit Base:</b> {info['bateria_padrão']}<br>
                    <b>📍 GPS / GNSS:</b> {info['gps']}<br>
                    <b>📶 Bluetooth:</b> {info['bluetooth']}<br>
                    <b>⚡ Potência de Saída:</b> {info['potencia']}<br>
                    <b>🚨 Man Down (Alarme de Queda):</b> {info['mandown']}
                </div>
            </div>
            """, unsafe_allow_html=True)
            
            col_check, col_qtd, col_bat = st.columns([2, 1, 1])
            with col_check:
                selecionado = st.checkbox(f"Incluir {info['nome']} no projeto", key=f"sel_{cat_key}")
            
            if selecionado:
                with col_qtd:
                    qtd = st.number_input(f"Quantidade de Rádios", min_value=1, value=1, key=f"qtd_{cat_key}")
                
                qtd_bat_extra = 0
                if info['tipo'] != 'Móvel / Veicular / Fixo':
                    with col_bat:
                        incluir_bat_extra = st.checkbox("🔋 Incluir Bateria Extra?", key=f"bat_extra_check_{cat_key}")
                        if incluir_bat_extra:
                            qtd_bat_extra = st.number_input("Qtd Baterias Extras", min_value=1, value=qtd, key=f"qtd_bat_extra_{cat_key}")
                
                carrinho_radios.append({
                    'categoria': cat_key,
                    'nome': info['nome'],
                    'qtd': qtd,
                    'qtd_bat_extra': qtd_bat_extra,
                    'compat_acessorio': info['compatibilidade_acessorio'],
                    'e_portatil': info['tipo'] != 'Móvel / Veicular / Fixo'
                })
            st.divider()

# --- 2. LICENÇAS DE SOFTWARE ---
st.header("2. Licenças de Software (SW Features)")
st.info("Escolha os recursos avançados de software para ativar nos rádios portáteis e/ou móveis selecionados.")

col_sw_port, col_sw_mov = st.columns(2)
has_portatil = any(r['e_portatil'] for r in carrinho_radios)
has_movel = any(not r['e_portatil'] for r in carrinho_radios)

with col_sw_port:
    st.subheader("💻 Licenças para Portáteis")
    if has_portatil:
        df_sw_port = df_lpu[df_lpu['Categoria'] == 'Portátil - SW FEATURES']
        for _, row in df_sw_port.iterrows():
            desc = row['Descritivo']
            if "TOGGLE RF POWER" in str(desc).upper():
                desc += " (Aumento de potência de emissão para 2.8W)"
            
            sw_sel = st.checkbox(f"{desc}", key=f"sw_port_{row['PN']}")
            if sw_sel:
                carrinho_sw.append({'Tipo': 'Portátil', 'PN': row['PN'], 'Descritivo': desc})
    else:
        st.caption("Selecione um rádio portátil na etapa 1 para ativar esta seção.")

with col_sw_mov:
    st.subheader("🚘 Licenças para Móveis / Veiculares")
    if has_movel:
        df_sw_mov = df_lpu[df_lpu['Categoria'] == 'Móvel - SW FEATURES']
        for _, row in df_sw_mov.iterrows():
            desc = row['Descritivo']
            sw_sel = st.checkbox(f"{desc}", key=f"sw_mov_{row['PN']}")
            if sw_sel:
                carrinho_sw.append({'Tipo': 'Móvel', 'PN': row['PN'], 'Descritivo': desc})
    else:
        st.caption("Selecione um rádio móvel na etapa 1 para ativar esta seção.")

st.divider()

# --- 3. ACESSÓRIOS ---
st.header("3. Acessórios Compatíveis")

if carrinho_radios:
    familias_compativeis = set([r['compat_acessorio'] for r in carrinho_radios])
    df_acessorios = df_lpu[df_lpu['Categoria'] == 'Portátil Acessórios'].copy()
    
    acessorios_disponiveis = []
    for _, row in df_acessorios.iterrows():
        grupo_str = str(row['Grupo'])
        for fam in familias_compativeis:
            if fam in grupo_str:
                acessorios_disponiveis.append(row)
                break
                
    if acessorios_disponiveis:
        df_acess_filtrados = pd.DataFrame(acessorios_disponiveis)
        st.write("Selecione os fones de ouvido, carregadores múltiplos e acessórios adicionais:")
        
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

st.divider()

# --- 4. RESUMO E FORMULÁRIO DE CONTATO ---
st.header("4. Resumo Geral da Configuração")

if carrinho_radios:
    st.subheader("📻 Rádios Escolhidos")
    resumo_r = []
    for r in carrinho_radios:
        item = {'Modelo / Terminal': r['nome'], 'Quantidade de Rádios': r['qtd']}
        if r['qtd_bat_extra'] > 0:
            item['Baterias Extras Solicitadas'] = r['qtd_bat_extra']
        else:
            item['Baterias Extras Solicitadas'] = "Nenhuma"
        resumo_r.append(item)
    st.table(pd.DataFrame(resumo_r))

    if carrinho_sw:
        st.subheader("💻 Licenças de Software Incluídas")
        st.table(pd.DataFrame(carrinho_sw)[['Tipo', 'Descritivo']])
        
    if carrinho_acessorios:
        st.subheader("🎧 Acessórios Adicionais")
        st.table(pd.DataFrame(carrinho_acessorios)[['Descritivo', 'Qtd']])

    st.success("Configuração concluída! Preencha seus dados de contato para gerar e enviar a solicitação.")
    
    with st.form("form_proposta"):
        c1, c2 = st.columns(2)
        with c1:
            nome_cliente = st.text_input("Nome do Solicitante *")
            empresa_cliente = st.text_input("Nome da Empresa / Órgão *")
        with c2:
            email_cliente = st.text_input("E-mail Corporativo *")
            telefone_cliente = st.text_input("Telefone de Contato com DDD *")
            
        submitted = st.form_submit_button("📩 Enviar Solicitação de Cotação")
        
        if submitted:
            if not nome_cliente or not email_cliente or not empresa_cliente or not telefone_cliente:
                st.error("Por favor, preencha todos os campos obrigatórios (*).")
            else:
                dados_c = {
                    'nome': nome_cliente,
                    'email': email_cliente,
                    'empresa': empresa_cliente,
                    'telefone': telefone_cliente
                }
                
                # Envio do e-mail
                sucesso, msg_status = enviar_email_cotacao(
                    destinatario="fabio.ashikaga@motorolasolutions.com",
                    dados_cliente=dados_c,
                    resumo_radios=resumo_r,
                    resumo_sw=carrinho_sw,
                    resumo_acc=carrinho_acessorios
                )
                
                if sucesso:
                    st.balloons()
                    st.success("Sua solicitação de cotação foi enviada com sucesso para fabio.ashikaga@motorolasolutions.com!")
                else:
                    st.warning(f"Solicitação registrada na tela, porém o e-mail automático não foi disparado: {msg_status}")