import streamlit as st
import pandas as pd
import platform

st.set_page_config(page_title="RPA: Proposta de acordos")

st.title('RPA: Proposta de acordos')
st.subheader("Coloque a planilha desejada, no formato definido:")

# Verifica se o sistema operacional é Windows e tenta importar as bibliotecas
try:
    if platform.system() == "Windows":
        import win32com.client as win32
        import pythoncom
        is_windows = True
    else:
        is_windows = False
except ImportError:
    st.error('Erro ao importar win32com.client e pythoncom. Certifique-se de que as bibliotecas estão instaladas corretamente.')
    is_windows = False

def enviar_email(df):
    if not is_windows:
        st.warning('O envio de emails só é suportado no Windows.')
        import win32com.client as win32
        import pythoncom
        return

    pythoncom.CoInitialize()

    try:
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
                    # Instanciando o Outlook
                    outlook = win32.Dispatch('outlook.application')
                    email = outlook.CreateItem(0)
                    # Configurando as informações do email
                    email.To = destinatario
                    email.Subject = f'{estado} - PROPOSTA DE ACORDO – {cnj} – {parte_contraria}'
                    # HTMLBody com a assinatura integrada
                    email.HTMLBody = f"""
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
                    email.Send()
                    st.success(f'Email enviado com sucesso para {destinatario}.')
                except Exception as e:
                    st.error(f'Erro ao enviar email para {destinatario}: {e}')
                    st.write(f'Detalhes do erro: {str(e)}')
            else:
                st.warning('Por favor, verifique os destinatários')

    except Exception as e:
        st.error(f'Erro na preparação do envio dos emails: {e}')
        st.write(f'Detalhes do erro: {str(e)}')
    finally:
        pythoncom.CoUninitialize()

# Upload do arquivo
df_file = st.file_uploader('Arraste aqui o relatório de acordo!', type=['csv', 'xlsx'])
if df_file is not None:
    df = pd.read_excel(df_file)
    st.dataframe(df.head(10))

    if st.button('Enviar'):
        enviar_email(df)
else:
    st.warning('Sem arquivo!')
