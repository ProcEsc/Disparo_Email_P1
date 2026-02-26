import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import smtplib
from email.message import EmailMessage
import os
import subprocess # Para rodar o LibreOffice no Linux

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações de E-mail")
    # Usando st.secrets para segurança ou inputs manuais
    email_usuario = st.text_input("E-mail Remetente:", placeholder="exemplo@empresa.com")
    senha_usuario = st.text_input("Senha de App:", type="password")

# --- FUNÇÃO DE CONVERSÃO LINUX (LIBREOFFICE) ---
def converter_para_pdf(caminho_docx):
    # O comando '--headless' roda o LibreOffice sem abrir interface gráfica
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf', 
            '--outdir', os.path.dirname(caminho_docx), 
            caminho_docx
        ], check=True)
        return caminho_docx.replace(".docx", ".pdf")
    except Exception as e:
        st.error(f"Erro na conversão PDF: {e}")
        return None

# --- FUNÇÃO DE ENVIO ---
def enviar_email_smtp(destinatario, nome_aluno, caminho_pdf, remetente, senha):
    msg = EmailMessage()
    msg['Subject'] = f"Boletim Escolar - {nome_aluno}"
    msg['From'] = remetente
    msg['To'] = destinatario
    msg.set_content(f"Olá {nome_aluno}, seu boletim segue em anexo.")

    with open(caminho_pdf, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='pdf', filename=os.path.basename(caminho_pdf))

    # Servidor Outlook/Office365
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)

# --- INTERFACE ---
st.title("🚀 Disparador de Boletins (Cloud Edition)")

col1, col2 = st.columns(2)
with col1:
    arq_excel = st.file_uploader("Upload Excel", type="xlsx")
with col2:
    arq_word = st.file_uploader("Upload Template Word", type="docx")

if arq_excel and arq_word:
    df = pd.read_excel(arq_excel)
    
    if st.button("🚀 Iniciar Processamento"):
        if not email_usuario or not senha_usuario:
            st.warning("Preencha as credenciais na barra lateral!")
        else:
            progresso = st.progress(0)
            for i, (idx, row) in enumerate(df.iterrows()):
                try:
                    # 1. Gerar Word
                    doc = DocxTemplate(arq_word)
                    doc.render(row.to_dict())
                    temp_docx = f"boletim_{row['Inscrição']}.docx"
                    doc.save(temp_docx)
                    
                    # 2. Converter PDF (Método Linux)
                    temp_pdf = converter_para_pdf(temp_docx)
                    
                    if temp_pdf and os.path.exists(temp_pdf):
                        # 3. Enviar
                        enviar_email_smtp(row['E-mail p4ed'], row['Nome'], temp_pdf, email_usuario, senha_usuario)
                        st.success(f"Enviado: {row['Nome']}")
                        
                        # Limpeza
                        os.remove(temp_docx)
                        os.remove(temp_pdf)
                except Exception as e:
                    st.error(f"Falha em {row['Nome']}: {e}")
                
                progresso.progress((i + 1) / len(df))
            st.balloons()