import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

st.set_page_config(page_title="RPA: Proposta de acordos")

st.title('RPA: Proposta de acordos')
st.subheader("Coloque a planilha desejada, no formato definido:")

def enviar_email(df):
    smtp_server = "smtp.office365.com"  # Servidor SMTP do Outlook
    smtp_port = 587  # Porta padrão para TLS
    smtp_user = "pm@nantesmello.com"  # Seu endereço de email do Outlook
    smtp_password = "Sodexo31"  # Sua senha de email

    try:
        # Conexão ao servidor SMTP
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)

        infos = {}
        for indice, linhas in df.iterrows():
            dicionario_linhas = linhas.to_dict()
            infos[indice] = dicionario_linhas

        i = 0
        for dicts in infos:
            destinatario = infos[i]['Destinatário']
            nome = infos[i]['Parte Contrária'].upper()
            cnj = infos[i]['Número do Processo'].upper()
            estado = infos[i]['Estado'].upper()
            parte_contraria = infos[i]['Parte Contrária'].upper()
            proposta = infos[i]['Proposta']

            i += 1
            confirmacao = '1'
            if confirmacao == '1':
                try:
                    # Configurando as informações do email
                    msg = MIMEMultipart()
                    msg['From'] = smtp_user
                    msg['To'] = destinatario
                    msg['Subject'] = f'{estado} - PROPOSTA DE ACORDO – {cnj} – {parte_contraria}'

                    # Corpo do email
                    html = f"""
                    <html>
                    <body>
                        <p>Prezado(a) {nome},</p>
                        <p>Segue a proposta de acordo para o processo {cnj}.</p>
                        <p>Atenciosamente,</p>
                        <p>{parte_contraria}</p>
                        <table width="520" border="0">
                            <tr>
                                <td width="480" align="center">
                                    <a href="https://www.nantesmello.com/" target="_blank">
                                        <img src="https://colosseo.com.br/email/nantes/logo.png" alt="Logo NM" width="137" height="94">
                                    </a>
                                </td>
                                <td>
                                    <table border="0">
                                        <tr>
                                            <td>
                                                <a href="mailto:pm@nantesmello.com">
                                                    <img src="https://colosseo.com.br/email/nantes/davidson-galdino-infos.jpg" alt="David Galdino">
                                                </a>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </body>
                    </html>
                    """
                    msg.attach(MIMEText(html, 'html'))

                    # Envio do email
                    server.sendmail(smtp_user, destinatario, msg.as_string())
                    st.success(f'Email enviado com sucesso para {destinatario}.')
                except Exception as e:
                    st.error(f'Erro ao enviar email para {destinatario}: {e}')
                    st.write(f'Detalhes do erro: {str(e)}')
            else:
                st.warning('Por favor, verifique os destinatários')

        server.quit()

    except Exception as e:
        st.error(f'Erro na preparação do envio dos emails: {e}')
        st.write(f'Detalhes do erro: {str(e)}')

# Upload do arquivo
df_file = st.file_uploader('Arraste aqui o relatório de acordo!', type=['csv', 'xlsx'])
if df_file is not None:
    df = pd.read_excel(df_file)
    st.dataframe(df.head(10))

    if st.button('Enviar'):
        enviar_email(df)
else:
    st.warning('Sem arquivo!')
