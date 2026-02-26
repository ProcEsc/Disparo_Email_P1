import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import smtplib
from email.message import EmailMessage
import os
from docx2pdf import convert

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações de E-mail")
    email_usuario = st.text_input("Seu e-mail (Remetente):", placeholder="exemplo@sistemapoliedro.com.br")
    senha_usuario = st.text_input("Senha de App:", type="password")

# --- FUNÇÃO DE ENVIO ---
def enviar_email_smtp(destinatario, nome_aluno, caminho_pdf, remetente, senha):
    msg = EmailMessage()
    msg['Subject'] = f"Boletim Escolar - {nome_aluno}"
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content(f"Olá {nome_aluno}, seu boletim está disponível em anexo.")

    with open(caminho_pdf, 'rb') as f:
        file_data = f.read()
        msg.add_attachment(
            file_data, 
            maintype='application', 
            subtype='pdf', 
            filename=os.path.basename(caminho_pdf)
        )

    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)

# --- INTERFACE PRINCIPAL ---
st.title("🚀 Disparador de Boletins")

col1, col2 = st.columns(2)
with col1:
    arq_excel = st.file_uploader("Upload da Planilha (Excel)", type="xlsx")
with col2:
    arq_word = st.file_uploader("Upload do Modelo (Word)", type="docx")

if arq_excel and arq_word:
    df = pd.read_excel(arq_excel)
    st.write(f"Total de alunos encontrados: {len(df)}")
    
    if st.button("🚀 Iniciar Disparos"):
        if not email_usuario or not senha_usuario:
            st.error("Por favor, preencha as credenciais na barra lateral.")
        else:
            progresso = st.progress(0)
            status_text = st.empty()
            
            for i, (index, row) in enumerate(df.iterrows()):
                nome_aluno = row['Nome']
                status_text.text(f"Processando: {nome_aluno}")
                
                try:
                    # 1. Gerar Word
                    doc = DocxTemplate(arq_word)
                    doc.render(row.to_dict())
                    nome_docx = f"temp_{row['Inscrição']}.docx"
                    doc.save(nome_docx)
                    
                    # 2. Converter para PDF
                    nome_pdf = nome_docx.replace(".docx", ".pdf")
                    convert(nome_docx, nome_pdf) 
                    
                    # 3. Enviar E-mail (Agora com todos os 5 argumentos necessários)
                    enviar_email_smtp(row['E-mail p4ed'], nome_aluno, nome_pdf, email_usuario, senha_usuario)
                    
                    st.toast(f"✅ Enviado: {nome_aluno}")
                    
                    # 4. Limpeza
                    os.remove(nome_docx)
                    os.remove(nome_pdf)
                    
                except Exception as e:
                    st.error(f"❌ Erro em {nome_aluno}: {e}")
                
                # Atualiza barra de progresso
                progresso.progress((i + 1) / len(df))
            
            status_text.text("Concluído!")
            st.balloons()