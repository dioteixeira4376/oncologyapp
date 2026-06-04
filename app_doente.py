#!/usr/bin/env python3
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import smtplib
from email.message import EmailMessage

st.set_page_config(
    page_title="Triagem Oncológica - Doente",
    page_icon="🩺",
    layout="wide"
)

st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #f7fbfc 0%, #eef8fa 100%);
}

.block-container {
    max-width: 1100px;
    padding-top: 2rem;
}

h1, h2, h3 {
    color: #0B2545;
}

.card {
    background: white;
    padding: 24px;
    border-radius: 18px;
    box-shadow: 0 8px 24px rgba(15, 23, 42, 0.06);
    margin-bottom: 20px;
}

.alert {
    background: #fff3f3;
    border-left: 6px solid #d90429;
    padding: 18px;
    border-radius: 12px;
    color: #7a0015;
    font-weight: 600;
}

.successbox {
    background: #edfdf5;
    border-left: 6px solid #00a86b;
    padding: 18px;
    border-radius: 12px;
    color: #064e3b;
    font-weight: 600;
}

div[data-testid="stButton"] button,
div[data-testid="stDownloadButton"] button {
    border-radius: 14px;
    background: linear-gradient(135deg, #0097A7, #006D77);
    color: white !important;
    border: none;
    padding: 0.65rem 1rem;
    font-weight: 600;
}
</style>
""", unsafe_allow_html=True)


QUEIXAS = [
    "Febre",
    "Náuseas",
    "Vómitos",
    "Diarreia",
    "Obstipação",
    "Mucosite",
    "Dor",
    "Fadiga",
    "Falta de ar",
    "Hemorragia",
    "Erupções cutâneas",
    "Prurido",
    "Neuropatia periférica",
    "Tonturas",
    "Perda de apetite",
    "Outros"
]

ICONES = {
    "Febre": "🌡️",
    "Náuseas": "🤢",
    "Vómitos": "🤮",
    "Diarreia": "🚽",
    "Obstipação": "🚽",
    "Mucosite": "👄",
    "Dor": "⚠️",
    "Fadiga": "🔋",
    "Falta de ar": "🫁",
    "Hemorragia": "🩸",
    "Erupções cutâneas": "🖐️",
    "Prurido": "🖐️",
    "Neuropatia periférica": "⚡",
    "Tonturas": "💫",
    "Perda de apetite": "🍽️",
    "Outros": "➕"
}


def enviar_email_alerta(dados_doente, queixas_graves):
    """
    Envia email se estiver configurado no Streamlit Secrets.

    No Streamlit Cloud, criar em Settings > Secrets:

    EMAIL_USER="teu_email@gmail.com"
    EMAIL_PASSWORD="app_password"
    EMAIL_TO="email_enfermagem@hospital.pt"
    SMTP_SERVER="smtp.gmail.com"
    SMTP_PORT=587
    """

    try:
        email_user = st.secrets["EMAIL_USER"]
        email_password = st.secrets["EMAIL_PASSWORD"]
        email_to = st.secrets["EMAIL_TO"]
        smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
    except Exception:
        return False, "Email não configurado em st.secrets."

    assunto = "⚠️ Alerta de Triagem Oncológica - Gravidade ≥ 3"

    linhas_queixas = "\n".join(
        [f"- {q['queixa']} | Gravidade {q['gravidade']} | Observação: {q['observacao']}" for q in queixas_graves]
    )

    corpo = f"""
Foi submetido um registo de triagem com gravidade ≥ 3.

Doente: {dados_doente['nome']}
Data de nascimento: {dados_doente['data_nascimento']}
Processo: {dados_doente['processo']}
Contacto: {dados_doente['contacto']}
Protocolo: {dados_doente['protocolo']}
Último tratamento: {dados_doente['ultimo_tratamento']}

Queixas graves:
{linhas_queixas}

Data/hora da submissão: {dados_doente['data_submissao']}

Contactar o doente com prioridade.
"""

    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = email_user
    msg["To"] = email_to
    msg.set_content(corpo)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as smtp:
            smtp.starttls()
            smtp.login(email_user, email_password)
            smtp.send_message(msg)

        return True, "Email enviado com sucesso."
    except Exception as e:
        return False, f"Erro ao enviar email: {e}"


def gerar_excel(registos):
    df = pd.DataFrame(registos)
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Registos")

    output.seek(0)
    return output


def main():
    with st.sidebar:
        st.markdown("## 🩺 Oncology App")
        st.markdown("---")
        st.markdown("🏠 Registo em casa")
        st.markdown("⚠️ Alerta automático")
        st.markdown("📋 Triagem de sintomas")
        st.markdown("---")
        st.caption("Esta ferramenta não substitui avaliação médica.")

    st.caption("TRIAGEM ONCOLÓGICA REMOTA")
    st.title("Registo de Queixas pelo Doente")
    st.write(
        "Preencha os seus sintomas. Se alguma queixa for graduada como 3 ou 4, "
        "a equipa de enfermagem será alertada para contacto."
    )

    if "registos" not in st.session_state:
        st.session_state.registos = []

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Identificação")

    col1, col2 = st.columns(2)

    with col1:
        nome = st.text_input("Nome completo")
        data_nascimento = st.date_input("Data de nascimento")
        processo = st.text_input("Número de processo")

    with col2:
        contacto = st.text_input("Contacto telefónico")
        protocolo = st.text_input("Protocolo / tratamento")
        ultimo_tratamento = st.date_input("Data do último tratamento")

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Queixas e gravidade")

    st.info("0 = sem sintomas | 1 = ligeiro | 2 = moderado | 3 = grave | 4 = muito grave")

    queixas_registadas = []
    col_a, col_b = st.columns(2)

    for idx, queixa in enumerate(QUEIXAS):
        coluna = col_a if idx % 2 == 0 else col_b

        with coluna:
            icone = ICONES.get(queixa, "➕")

            selecionada = st.checkbox(f"{icone} {queixa}", key=f"check_{queixa}")

            if selecionada:
                nome_queixa = queixa

                if queixa == "Outros":
                    nome_queixa = st.text_input("Descreva a outra queixa")

                gravidade = st.slider(
                    f"Gravidade - {queixa}",
                    min_value=0,
                    max_value=4,
                    value=1,
                    step=1,
                    key=f"gravidade_{queixa}"
                )

                observacao = st.text_area(
                    f"Observações - {queixa}",
                    key=f"obs_{queixa}",
                    height=80
                )

                if nome_queixa:
                    queixas_registadas.append(
                        {
                            "queixa": nome_queixa,
                            "gravidade": gravidade,
                            "observacao": observacao
                        }
                    )

    st.markdown("</div>", unsafe_allow_html=True)

    consentimento = st.checkbox(
        "Confirmo que estes dados podem ser enviados à equipa de enfermagem para efeitos de triagem/contacto."
    )

    submitted = st.button("Submeter sintomas")

    if submitted:
        if not nome or not contacto or not processo:
            st.error("Preencha pelo menos nome, contacto e número de processo.")
            return

        if not consentimento:
            st.error("É necessário confirmar o consentimento antes de submeter.")
            return

        if len(queixas_registadas) == 0:
            st.error("Selecione pelo menos uma queixa.")
            return

        data_submissao = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        dados_doente = {
            "nome": nome,
            "data_nascimento": str(data_nascimento),
            "processo": processo,
            "contacto": contacto,
            "protocolo": protocolo,
            "ultimo_tratamento": str(ultimo_tratamento),
            "data_submissao": data_submissao,
        }

        queixas_graves = [
            q for q in queixas_registadas if q["gravidade"] >= 3
        ]

        for q in queixas_registadas:
            st.session_state.registos.append(
                {
                    "Data submissão": data_submissao,
                    "Nome": nome,
                    "Data nascimento": str(data_nascimento),
                    "Processo": processo,
                    "Contacto": contacto,
                    "Protocolo": protocolo,
                    "Último tratamento": str(ultimo_tratamento),
                    "Queixa": q["queixa"],
                    "Gravidade": q["gravidade"],
                    "Observação": q["observacao"],
                    "Alerta": "Sim" if q["gravidade"] >= 3 else "Não"
                }
            )

        if queixas_graves:
            enviado, mensagem = enviar_email_alerta(dados_doente, queixas_graves)

            st.markdown(
                '<div class="alert">⚠️ Foram registadas queixas com gravidade 3 ou 4. '
                'A equipa de enfermagem deve contactar o doente com prioridade.</div>',
                unsafe_allow_html=True
            )

            if enviado:
                st.success(mensagem)
            else:
                st.warning(
                    "Registo efetuado, mas o email de alerta ainda não está configurado. "
                    f"Detalhe: {mensagem}"
                )
        else:
            st.markdown(
                '<div class="successbox">✅ Registo submetido com sucesso. '
                'As queixas não atingem o limiar de alerta automático.</div>',
                unsafe_allow_html=True
            )

    st.markdown("---")
    st.subheader("Área de enfermagem - registos desta sessão")

    if st.session_state.registos:
        df = pd.DataFrame(st.session_state.registos)
        st.dataframe(df, use_container_width=True)

        excel_file = gerar_excel(st.session_state.registos)

        st.download_button(
            "Download dos registos em Excel",
            data=excel_file,
            file_name="registos_triagem_doente.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.caption("Ainda não existem registos nesta sessão.")


if __name__ == "__main__":
    main()
