import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import smtplib
from email.message import EmailMessage
import os
import subprocess
import tempfile 
import shutil


with st.sidebar:
    st.header("Configurações de E-mail")
    # Tenta pegar dos Secrets, se não existir, pede no input
    email_usuario = st.text_input("E-mail Remetente:")
    senha_usuario = st.text_input("Senha / App Password:", type="password")

# --- FUNÇÃO DE CONVERSÃO ---
def converter_para_pdf(caminho_docx, pasta_saida):
    try:
        subprocess.run([
            'libreoffice', '--headless', '--convert-to', 'pdf', 
            '--outdir', pasta_saida, 
            caminho_docx
        ], check=True, capture_output=True) # capture_output evita poluir logs
        
        nome_pdf = os.path.basename(caminho_docx).replace(".docx", ".pdf")
        return os.path.join(pasta_saida, nome_pdf)
    except Exception as e:
        st.error(f"Erro na conversão PDF: {e}")
        return None

# --- FUNÇÃO DE ENVIO ---
def enviar_email_smtp(destinatarios, nome_aluno, caminho_pdf, remetente, senha):
    msg = EmailMessage()
    msg['Subject'] = f"Boletim Escolar - {nome_aluno}"
    msg['From'] = remetente
    msg['To'] = destinatarios 
    msg.set_content(f"Olá, {nome_aluno}! As notas da sua P1 do primeiro trimestre seguem em anexo. Confira o gabarito e resolução disponibilizadas em seu drive! ")

    with open(caminho_pdf, 'rb') as f:
        msg.add_attachment(
            f.read(), 
            maintype='application', 
            subtype='pdf', 
            filename=os.path.basename(caminho_pdf)
        )

    # Use port 587 para STARTTLS
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)

# --- INTERFACE ---
st.title("Disparador de Boletins")

arq_excel = st.file_uploader("Upload Planilha (Excel)", type="xlsx")
arq_word = st.file_uploader("Upload Modelo (Word)", type="docx")

if arq_excel and arq_word:
    df = pd.read_excel(arq_excel)
    st.info(f"📋 {len(df)} registros encontrados.")
    
    if st.button("Iniciar Processamento"):
        if not email_usuario or not senha_usuario:
            st.warning("Credenciais ausentes!")
        else:
            progresso = st.progress(0)
            
            for i, (idx, row) in enumerate(df.iterrows()):
                nome_aluno = row.get('Nome', 'Aluno')
                
                # CRIAR PASTA TEMPORÁRIA ÚNICA POR ALUNO/SESSÃO
                # Isso impede que um usuário veja dados de outro
                with tempfile.TemporaryDirectory() as pasta_tmp:
                    try:
                        # 1. Gerar Word dentro da pasta temporária
                        doc = DocxTemplate(arq_word)
                        doc.render(row.to_dict())
                        
                        temp_docx = os.path.join(pasta_tmp, f"boletim_{row['Inscrição']}.docx")
                        doc.save(temp_docx)
                        
                        # 2. Converter (PDF será salvo na mesma pasta tmp)
                        temp_pdf = converter_para_pdf(temp_docx, pasta_tmp)
                        
                        if temp_pdf and os.path.exists(temp_pdf):
                            # 3. Validar E-mails
                            lista_bruta = [row.get('E-mail p4ed'), row.get('E-mail RP Pessoal')]
                            emails_validos = [str(e).strip() for e in lista_bruta if pd.notna(e) and "@" in str(e)]
                            
                            if emails_validos:
                                destinatarios_str = ", ".join(emails_validos)
                                enviar_email_smtp(destinatarios_str, nome_aluno, temp_pdf, email_usuario, senha_usuario)
                                st.toast(f"✅ Enviado: {nome_aluno}")
                            else:
                                st.warning(f"⚠️ Sem e-mail válido: {nome_aluno}")
                                
                    except Exception as e:
                        st.error(f"❌ Erro em {nome_aluno}: {e}")
                    
                    # AO SAIR DO 'WITH', A PASTA TEMPORÁRIA É EXCLUÍDA AUTOMATICAMENTE
                    # Mesmo que ocorra um erro catastrófico acima.
                
                progresso.progress((i + 1) / len(df))
            
            st.success("✅ Processo finalizado!")