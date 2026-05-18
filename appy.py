#!/usr/bin/env python3
import streamlit as st
import pandas as pd
import base64
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side

# Configura a página para o layout centralizado
st.set_page_config(
    page_title="Oncology App",
    page_icon="logo.png",
    layout="wide"
)
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #f7fbfc 0%, #eef8fa 100%);
}

.block-container {
    padding-top: 2rem;
    max-width: 1200px;
}

h1, h2, h3 {
    color: #0B2545;
}

section[data-testid="stSidebar"] {
    background: #ffffff;
    border-right: 1px solid #e6eef2;
}

div[data-testid="stButton"] button,
div[data-testid="stLinkButton"] a,
div[data-testid="stDownloadButton"] button {
    border-radius: 14px;
    background: linear-gradient(135deg, #0097A7, #006D77);
    color: white !important;
    border: none;
    padding: 0.65rem 1rem;
    font-weight: 600;
    box-shadow: 0 8px 18px rgba(0, 109, 119, 0.18);
}

div[data-testid="stButton"] button:hover,
div[data-testid="stLinkButton"] a:hover {
    transform: translateY(-1px);
    box-shadow: 0 10px 22px rgba(0, 109, 119, 0.24);
}

div[data-testid="stCheckbox"] {
    background: white;
    padding: 0.65rem 0.9rem;
    border-radius: 14px;
    border: 1px solid #e6eef2;
    margin-bottom: 0.45rem;
    box-shadow: 0 4px 14px rgba(15, 23, 42, 0.04);
}

div[data-testid="stCheckbox"]:hover {
    border-color: #00A6B4;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
}

.stSelectbox, .stTextInput, .stDateInput {
    background: white;
    border-radius: 14px;
}

[data-testid="stMetric"] {
    background: white;
    padding: 1rem;
    border-radius: 16px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
}
</style>
""", unsafe_allow_html=True)

# Definir as intervenções recomendadas para cada queixa
intervencoes = {
    "Febre": [
        "Monitorar temperatura e sintomas.",
        "Considerar antipiréticos e hidratação.",
        "Consultar médico, possíveis exames laboratoriais.",
        "Avaliação médica imediata, hospitalização possível."
    ],
    "Náuseas": [
        "Recomendar pequenas refeições frequentes, gengibre.",
        "Prescrever antieméticos, monitorar ingestão de líquidos.",
        "Avaliação médica, ajustar medicação antiemética.",
        "Hospitalização, suporte intravenoso."
    ],
    # Adicionar intervenções para outras queixas aqui...
}

# Definir classe Paciente
class Paciente:
    def __init__(self, nome, data_de_nascimento, processo, médico_assistente, protocolo, último_tratamento, queixas, intervencoes_selecionadas):
        self.nome = nome
        self.data_de_nascimento = data_de_nascimento
        self.processo = processo
        self.médico_assistente = médico_assistente
        self.protocolo = protocolo
        self.último_tratamento = último_tratamento
        self.queixas = queixas
        self.intervencoes_selecionadas = intervencoes_selecionadas

# Definir classe TriagemQuimioterapia
class TriagemQuimioterapia:
    def __init__(self):
        self.protocolo = ""
        self.pacientes = []

    def adicionar_paciente(self, paciente):
        self.pacientes.append(paciente)

    def realizar_triagem(self, paciente):
        dados_triagem = {
            "Nome do paciente": paciente.nome,
            "Data de nascimento": paciente.data_de_nascimento,
            "Processo": paciente.processo,
            "Médico assistente": paciente.médico_assistente,
            "Último tratamento": paciente.último_tratamento,
            "Protocolo de quimioterapia": paciente.protocolo,
        }
        # Adicionar queixas e intervenções em células separadas
        for i, queixa in enumerate(paciente.queixas):
            nome_queixa, _ = queixa.split(" - Gravidade ")
            dados_triagem[f"Queixa {i+1}"] = queixa
            intervencoes_para_queixa = [interv for queixa_sel, interv in paciente.intervencoes_selecionadas if queixa_sel == nome_queixa]
            dados_triagem[f"Intervenção para Queixa {i+1}"] = "; ".join(intervencoes_para_queixa)
        return dados_triagem

# Função para coletar dados do paciente
def coletar_dados_paciente():
    nome = st.text_input("Nome do paciente:")
    col1, col2 = st.columns(2)
    with col1:
        data_atual = datetime.now().date()
        data_inicial = datetime(1900, 1, 1).date()
        data_de_nascimento = st.date_input("Data de Nascimento:", min_value=data_inicial, max_value=data_atual)
    with col2:
        processo = st.text_input("Número do Processo")
    protocolo = ["AC", "Avelumab", "Bevacizumab+Xelox", "Bevacizumab+Folfiri", "Bevacizumab+Folfox", "Cabazitaxel",
                 "Cetuximab+Folfox", "Cetuximab+Gramont", "Ciclofosfamida", "DCF mod", "Folfox", "Folfiri",
                 "Nab - Paclitaxel+Gemcitabina", "Nivolumab", "Paclitaxel", "Pertuzumab+trastuzumab",
                 "Pertuzumab+trastuzumab+docetaxel", "Pembrolizumab", "Pertuzumab+Trastuzumab+Paclitaxel", "TPE",
                 "Trabectedina", "Trabectedina+Caelyx", "Trastuzumab emtansina", "Vincristina", "Atezolizumab",
                 "TCH (Docetaxel+Carboplatina+Trastuzumab)", "Docetaxel+Carboplatina", "Xelox+ Bevacizumab", "Xelox",
                 "ELF (Etoposido+Leucovorin+5Fu)", "Doxorrubicina", "Zometa / Ácido Zoledrónico", "Alimta", "Topotecano",
                 "Gemcitabina", "Vinorelbina", "Trastuzumab EV", "Irinotecano+Cetuximab", "Bleomicina",
                 "Etoposido+Carboplatina", "Gemcitabina+Carboplatina", "Carboplatina+Vinorrelbina oral",
                 "Paclitaxel+Gemcitabina", "Etoposido", "Cetuximab", "Irinotecano", "Caelyx", "Cetuximab + Paclitaxel",
                 "PCV", "Paclitaxel+carboplatina", "Paclitaxel+Cetuximab", "Panitumumab", "Panitumumab+Folfox",
                 "Paclitaxel+Ramucirumab", "Ramucirumab", "Pembrolizumab", "FLOT", "Trastuzumab Deruxtecano",
                 "TCHP (Pertuzumab+Trastuzumab+Docetaxel+Carboplatina)", "Cemiplimab", "Sacituzumab Govitecano",
                 "Pembrolizumab+Lenvatinib", "Cetuximab+Encorafenib", "Pembrolizumab+carboplatina", "Enfortumab",
                 "Trabectadina", "Folfnaliri", "5FU+RT", "Atezolizumab + bevacizumab", "Pembrolizumab+Cisplatina+5FU",
                 "Trastuzumab+Folfox", "PacCHP (Pert+Trast+Carbo+Paclitaxel)", "Carboplatina+5FU", "Trastuzumab+FLOT",
                 "MVAC (Metotrexato+Vimblastina+Doxorrubicina+Cisplatina)", "Bevacizumab+Gramont", "Avelumab",
                 "Paclitaxel+Carboplatina", "Cisplatina+Epirrubicina+5FU", "Cisplatina+Docetaxel", "Cisplatina+Etoposido",
                 "Cisplatina+5FU", "Cetuximab+Folfiri", "Docetaxel+Trastuzumab", "Cisplatina+Alimta",
                 "Cisplatina+VNB oral", "Cisplatina", "Paclitaxel+Gemcitabina", "Xelox", "TC (Docetaxel+Ciclofosfamida)",
                 "Gramont", "Cisplatina+Gencitabina", "Mitomicina+5FU", "Docetaxel+Epirrubicina", "BEP", "TPF",
                 "Cisplatina+Irinotecano", "Trastuzumab", "Paclitaxel+Carboplatina+Trastuzumab", "Vinorelbina+Gemcitabina",
                 "Gemcitabina+Trastuzumab", "Mayo", "Cetuximab+Xelox", "Taxotere/Docetaxel",
                 "Cisplatina+Bleomicina+Metotrexato", "Nab- paclitaxel", "Folfirinox", "Metotrexato",
                 "Panitumumab+Folfiri", "Panitumumab+Gramont", "Cetuximab+Cisplatina+5FU", "Mitoxantrone",
                 "Dacarbazina", "Folfirinox+Bevacizumab", "Nivolumab+Ipilimumab", "Cisplatina+Gemcitabina+Durvalumab",
                 "Durvalumab", "TC Phesgo (Docetaxel+Carboplatina+Phesgo)", "PacC Phesgo  (Paclitaxel+Carboplatina+Phesgo)",
                 "Paclitaxel+Phesgo", "Phesgo", "Cetuximab+Carboplatina+5FU", "Carboplatina", "Bevacizumab",
                 "Denosumab (SC)", "Vinorelbina", "5FU+RT"]
    protocolo_selecionado = st.selectbox("Selecione o protocolo:", protocolo)
    
    st.header("Queixa(s) referida(s)")
    st.link_button(
    "Abrir Documento de Triagem",
    "https://www.ukacuteoncology.co.uk/application/files/4217/3799/3595/UKONS_Triage_tool_poster_A3_V10.pdf"
)

    queixas = ["Alteração da função hepática", "Anemia", "Artralgia", "Alopecia", "Astenia", "Cefaleias",
               "Cardiomiopatia irreversível", "Cistite hemorrágica", "Dermatite", "Diarreia", "Disgeusia",
               "Disosmia", "Eritrodisestesia palmoplantar", "Erupções cutâneas", "Fadiga", "Febre", "Fibrose Pulmonar",
               "Hemorragia", "Insónia", "Mucosite", "Náuseas", "Neuropatia Central", "Neuropatia periférica",
               "Nefrotoxicidade", "Obstipação", "Ototoxicidade", "Paroniquia", "Perda de apetite",
               "Prurido", "Reações de hipersensibilidade", "Síndrome Colinérgico", "Síndrome de retenção de fluidos",
               "Toxicidade neurológica", "Toxicidade renal", "Tonturas", "Vómitos"]
    icones_queixas = {
    "Alteração da função hepática", "Anemia": "🩸", "Artralgia", "Alopecia": "👨🏻‍🦲", "Astenia", "Cefaleias": "🤕",
               "Cardiomiopatia irreversível": "🫀", "Cistite hemorrágica", "Dermatite", "Diarreia": "🚽", "Disgeusia",
               "Disosmia", "Eritrodisestesia palmoplantar": "✋🏽", "Erupções cutâneas", "Fadiga", "Febre": "🌡️", "Fibrose Pulmonar": "🫁",
               "Hemorragia": "🩸", "Insónia": "😴", "Mucosite": "👄", "Náuseas": "🤢", "Neuropatia Central": "🧠", "Neuropatia periférica": "⚡",
               "Nefrotoxicidade", "Obstipação": "🚽", "Ototoxicidade": "👂", "Paroniquia": "💅🏻", "Perda de apetite",
               "Prurido", "Reações de hipersensibilidade", "Síndrome Colinérgico", "Síndrome de retenção de fluidos",
               "Toxicidade neurológica", "Toxicidade renal", "Tonturas", "Vómitos": "🤮"
}

    # Lista completa de queixas
    queixas_gravidade = []

    # Dividindo as queixas em duas colunas
    col1, col2 = st.columns(2)
    mid_point = len(queixas) // 2

    with col1:
        for queixa in queixas[:mid_point]:
            icone = icones_queixas.get(queixa, "➕")
            if st.checkbox(f"{icone}  {queixa}", key=f"{queixa}_1"):
                gravidade = st.slider(f"Gravidade de {queixa} (CTCAE)", min_value=0, max_value=4, step=1, key=f"grav_{queixa}_1")
                queixas_gravidade.append(f"{queixa} - Gravidade {gravidade}")

    with col2:
        for queixa in queixas[mid_point:]:
            icone = icones_queixas.get(queixa, "➕")
            if st.checkbox(f"{icone}  {queixa}", key=f"{queixa}_2"):
                gravidade = st.slider(f"Gravidade de {queixa} (CTCAE)", min_value=0, max_value=4, step=1, key=f"grav_{queixa}_2")
                queixas_gravidade.append(f"{queixa} - Gravidade {gravidade}")

    # Entrada personalizada para "Outros"
    if st.checkbox("Outros", key="Outros"):
        outra_queixa = st.text_input("Digite a outra queixa:")
        if outra_queixa:
            gravidade = st.slider(f"Gravidade de {outra_queixa} (CTCAE)", min_value=0, max_value=4, step=1, key=f"grav_{outra_queixa}")
            queixas_gravidade.append(f"{outra_queixa} - Gravidade {gravidade}")

    intervencoes_selecionadas = []
    for queixa in queixas_gravidade:
        nome_queixa, _ = queixa.split(" - Gravidade ")
        st.subheader(f"Sugestões de Intervenção para {nome_queixa}")
        if nome_queixa in intervencoes:
            for intervencao in intervencoes.get(nome_queixa, []):
                if st.checkbox(intervencao, key=f"interv_{nome_queixa}_{intervencao}"):
                    intervencoes_selecionadas.append((nome_queixa, intervencao))
            # Adicionar opção para intervenção personalizada
            if st.checkbox("Outra(s)", key=f"interv_{nome_queixa}_outras"):
                intervencao_personalizada = st.text_input(f"Digite a intervenção para {nome_queixa}:", key=f"input_{nome_queixa}_outras")
                if intervencao_personalizada:
                    intervencoes_selecionadas.append((nome_queixa, intervencao_personalizada))
        else:
            intervencao_personalizada = st.text_input(f"Digite a intervenção para {nome_queixa}:")
            if intervencao_personalizada:
                intervencoes_selecionadas.append((nome_queixa, intervencao_personalizada))

    último_tratamento = st.date_input("Data de último tratamento:")
    médico_assistente = st.selectbox("Selecione o médico assistente:", ["Sandra Custódio", "Joana Marinho", "Enrique Dias", "Moreira Pinto", "Andreia Capela", "Cristiana Marques", "Inês Leão",
        "Sandra Silva", "Helena Guedes", "Adriana Soares", "Ana Raquel Monteiro", "Raquel Basto"])  # Exemplo de opções

    return Paciente(nome, data_de_nascimento, processo, médico_assistente, protocolo_selecionado, último_tratamento, queixas_gravidade, intervencoes_selecionadas)

def mostrar_legenda():
    st.markdown("""
    <style>
    .square {
        height: 25px;
        width: 25px;
        display: inline-block;
    }
    .lightgreen { background-color: lightgreen; }
    .yellow { background-color: yellow; }
    .orange { background-color: orange; }
    .red { background-color: red; }
    .column {
        float: left;
        width: 50%;  /* Define cada coluna para usar 50% da largura disponível */
        padding: 5px; /* Opcional: adicionar um pouco de espaço */
    }
    .row:after {
        content: "";
        display: table;
        clear: both;
    }
    </style>
    <div class="row">
        <div class="column">
            <span class='square lightgreen'></span> Gravidade 1 (conselhos sobre autocuidado) <br>
            <span class='square yellow'></span> Gravidade 2 (reavaliação em até 24 horas) <br>
        </div>
        <div class="column">
            <span class='square orange'></span> Gravidade 3 (referir para serviço médico) <br>
            <span class='square red'></span> Gravidade 4 (compareça para avaliação o mais rápido possível)
        </div>
    </div>
    """, unsafe_allow_html=True)

def color_definition(val):
    if isinstance(val, str):
        if "Gravidade 0" in val:
            color = 'FFFFE0'  # lightyellow
        elif "Gravidade 1" in val:
            color = '90EE90'  # lightgreen
        elif "Gravidade 2" in val:
            color = 'FFFF00'  # yellow
        elif "Gravidade 3" in val:
            color = 'FFA500'  # orange
        elif "Gravidade 4" in val:
            color = 'FF0000'  # red
        else:
            color = 'FFFFFF'  # default white
        return color
    else:
        return 'FFFFFF'  # default white if val is not a string

# Função principal
st.markdown("""
<style>
.stApp {
    background: linear-gradient(180deg, #f7fbfc 0%, #eef6f8 100%);
}

h1, h2, h3 {
    color: #0B3C5D;
}

div[data-testid="stButton"] button {
    border-radius: 12px;
    border: none;
    background: #0B7893;
    color: white;
    font-weight: 600;
    padding: 0.6rem 1rem;
}

div[data-testid="stButton"] button:hover {
    background: #095f75;
    color: white;
}

div[data-testid="stDownloadButton"] button {
    border-radius: 12px;
    background: #0B7893;
    color: white;
}

section[data-testid="stSidebar"] {
    background-color: #ffffff;
}
</style>
""", unsafe_allow_html=True)
def main():

    with st.sidebar:
        st.image("https://www.ulsge.min-saude.pt/static/media/logo.a59fd095.png", width=130)
        st.markdown("## Oncology App")
        st.markdown("---")
        st.markdown("📋 Triagem de Pacientes")
        st.markdown("👤 Novo Paciente")
        st.markdown("📊 Estatísticas")
        st.markdown("ℹ️ Sobre a App")

    col_logo, col_title, col_btn = st.columns([1, 4, 2])

    with col_logo:
        st.image("https://www.ulsge.min-saude.pt/static/media/logo.a59fd095.png", width=90)

    with col_title:
        st.caption("TRIAGEM")
        st.title("Triagem de Pacientes")
        st.write("Selecione os sintomas apresentados pelo paciente.")

    with col_btn:
        st.link_button(
            "📄 Abrir Documento de Triagem",
            "https://www.ukacuteoncology.co.uk/wp-content/uploads/2024/10/UKONS-AOS-Triage-Tool-November-2024.pdf"
        )

    @st.cache(allow_output_mutation=True)
    def get_triagem():
        return TriagemQuimioterapia()

    triagem = get_triagem()

    st.header("Novo Paciente")
    novo_paciente = coletar_dados_paciente()
    if st.button("Adicionar Paciente"):
        triagem.adicionar_paciente(novo_paciente)
        st.success("Paciente adicionado com sucesso!")

    st.header("Visualizar Triagem de Pacientes")
    if len(triagem.pacientes) > 0:
        df = pd.DataFrame([triagem.realizar_triagem(paciente) for paciente in triagem.pacientes])
            
        # Aplicando cores com base na gravidade das queixas
        queixas_colunas = [col for col in df.columns if col not in ["Nome do paciente", "Data de nascimento", "Processo", "Médico assistente", "Último tratamento", "Protocolo de quimioterapia"]]
        df_styled = df.style.map(lambda x: f'background-color: #{color_definition(x)}; color: black;' if isinstance(x, str) else '', subset=pd.IndexSlice[:, queixas_colunas])
        st.dataframe(df_styled)

        # Botão para exportar para Excel com estilos e bordas
        if st.button("Exportar para Excel"):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Triagem')
                workbook = writer.book
                worksheet = writer.sheets['Triagem']
                
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                for col in worksheet.columns:
                    for cell in col:
                        cell_value = cell.value
                        color = color_definition(cell_value)
                        if color:
                            cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
                        cell.border = thin_border

            output.seek(0)
            excel_data = output.getvalue()
            b64 = base64.b64encode(excel_data).decode()
            href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="triagem_pacientes.xlsx">Clique aqui para baixar o arquivo</a>'
            st.markdown(href, unsafe_allow_html=True)
            st.success("Dados exportados para triagem_pacientes.xlsx com sucesso!")

    else:
        st.warning("Ainda não há pacientes adicionados.")
            
    mostrar_legenda()

# Executar a função principal
if __name__ == "__main__":
    main()
