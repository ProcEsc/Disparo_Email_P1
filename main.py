import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import smtplib
from email.message import EmailMessage
import os
import subprocess

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações de E-mail")
    email_usuario = st.text_input("E-mail Remetente:", placeholder="exemplo@empresa.com")
    senha_usuario = st.text_input("Senha:", type="password")

# --- FUNÇÃO DE CONVERSÃO (LINUX/LIBREOFFICE) ---
def converter_para_pdf(caminho_docx):
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
def enviar_email_smtp(destinatarios, nome_aluno, caminho_pdf, remetente, senha):
    msg = EmailMessage()
    msg['Subject'] = f"Boletim Escolar - {nome_aluno}"
    msg['From'] = remetente
    
    # O campo 'To' aceita uma string com e-mails separados por vírgula
    msg['To'] = destinatarios 
    
    msg.set_content(f"Olá {nome_aluno}! Seu boletim da P1 do primeiro trimestre segue em anexo.")

    with open(caminho_pdf, 'rb') as f:
        msg.add_attachment(
            f.read(), 
            maintype='application', 
            subtype='pdf', 
            filename=os.path.basename(caminho_pdf)
        )

    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.starttls()
        server.login(remetente, senha)
        server.send_message(msg)

# --- INTERFACE PRINCIPAL ---
st.title("Disparador de Boletins")

col1, col2 = st.columns(2)
with col1:
    arq_excel = st.file_uploader("Upload Planilha (Excel)", type="xlsx")
with col2:
    arq_word = st.file_uploader("Upload Modelo (Word)", type="docx")

if arq_excel and arq_word:
    df = pd.read_excel(arq_excel)
    st.write(f"Alunos identificados: {len(df)}")
    
    if st.button("🚀 Iniciar Processamento"):
        if not email_usuario or not senha_usuario:
            st.warning("Preencha as credenciais na barra lateral!")
        else:
            progresso = st.progress(0)
            status_text = st.empty()
            
            for i, (idx, row) in enumerate(df.iterrows()):
                nome_aluno = row['Nome']
                status_text.text(f"Processando: {nome_aluno}")
                
                try:
                    # 1. Gerar Word
                    doc = DocxTemplate(arq_word)
                    doc.render(row.to_dict())
                    temp_docx = f"Inscricao_{row['Inscrição']}.docx"
                    doc.save(temp_docx)
                    
                    # 2. Converter para PDF
                    temp_pdf = converter_para_pdf(temp_docx)
                    
                    if temp_pdf:
                        # 3. Tratamento e Consolidação de E-mails (Data Quality)
                        # Coleta os e-mails das colunas identificadas
                        lista_bruta = [row['E-mail p4ed'], row['E-mail RP Pessoal']]
                        
                        # Filtra valores vazios e limpa espaços
                        emails_validos = [str(e).strip() for e in lista_bruta if pd.notna(e) and "@" in str(e)]
                        
                        if emails_validos:
                            # Junta os e-mails em uma string separada por vírgula
                            destinatarios_str = ", ".join(emails_validos)
                            
                            # 4. Enviar E-mail
                            enviar_email_smtp(destinatarios_str, nome_aluno, temp_pdf, email_usuario, senha_usuario)
                            st.toast(f"✅ Enviado: {nome_aluno}")
                        else:
                            st.error(f"❌ Nenhum e-mail válido para {nome_aluno}")
                        
                        # 5. Limpeza de arquivos temporários
                        os.remove(temp_docx)
                        os.remove(temp_pdf)
                        
                except Exception as e:
                    st.error(f"❌ Falha ao processar {nome_aluno}: {e}")
                
                progresso.progress((i + 1) / len(df))
            
            status_text.text("Concluído!")
            st.balloons()
            st.success("Processamento finalizado com sucesso!")