import streamlit as st
import pandas as pd
import platform
import pythoncom

st.set_page_config(page_title="RPA: Proposta de acordos")

st.title('RPA: Proposta de acordos')
st.subheader("Coloque a planilha desejada, no formato definido:")

# Verifica se o sistema operacional é Windows
is_windows = platform.system() == "Windows"

if is_windows:
    try:
        import win32com.client as win32
        import pythoncom
    except ImportError:
        is_windows = False

def enviar_email(df):
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
                    if is_windows:
                        # Instanciando o Outlook
                        outlook = win32.Dispatch('outlook.application')
                        email = outlook.CreateItem(0)
                        # Configurando as informações do email
                        email.To = destinatario
                        email.Subject = f'{estado} - PROPOSTA DE ACORDO – {cnj} – {parte_contraria}'
                        # HTMLBody com a assinatura integrada
                        email.HTMLBody = f"""
                                <b><span style='font-size:12.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Proposta de Acordo<o:p></o:p></span></b></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Prezado(a),<o:p></o:p></span></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Me chamo Davidson, faço parte do time do Nantes Mello Advogados e somos responsáveis pela condução dos processos judiciais envolvendo a Azul.<o:p></o:p></span></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Venho por meio deste e-mail realizar proposta de acordo visando a encerrar a lide de forma amigável e antecipada.<o:p></o:p></span></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><b><span style='font-size:12.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Benefícios da Proposta<o:p></o:p></span></b></p>
                                <ul type=disc>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo3;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Rapidez<b>:</b>&nbsp;Encerramento antecipado da lide poupando recursos e tempo para ambas as partes.<o:p></o:p></span></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo3;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Eficiência<b>:</b>&nbsp;Redução dos custos processuais e administrativos associados à continuidade do litígio.<o:p></o:p></span></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l0 level1 lfo3;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Respeito e Satisfação do Cliente<b>:</b>&nbsp;Possibilidade de encontrar uma solução mutuamente satisfatória preservando a relação entre as partes.<o:p></o:p></span></li>
                                </ul>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><b><span style='font-size:12.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Proposta<o:p></o:p></span></b></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>&#8220;A AZUL por mera liberalidade compromete-se a disponibilizar <span style='background:yellow;mso-highlight:yellow'>{proposta} voucher(s)</span> no prazo máximo de até 15 (quinze) dias úteis às partes mencionadas abaixo (devendo a parte providenciar desde já o adequado cadastro no programa de fidelidade TudoAzul):<o:p></o:p></span></i></p>
                                <ul type=disc>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l1 level1 lfo6;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Nome Completo (parte):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l1 level1 lfo6;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>CPF (parte):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l1 level1 lfo6;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Conta AZUL FIDELIDADE (parte):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l1 level1 lfo6;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>e-mail (parte):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l1 level1 lfo6;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Quantidade de voucher(s):<o:p></o:p></span></i></li>
                                </ul>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>(caso a emissão de voucher seja dividida entre Autor e Advogado incluir dados do advogado):<o:p></o:p></span></i></p>
                                <ul type=disc>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Nome Completo (patrono):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>CPF (patrono):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Conta AZUL FIDELIDADE (patrono):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>e-mail (patrono):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>OAB/UF (patrono):<o:p></o:p></span></i></li>
                                <li class=MsoNormal style='color:#242424;mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-list:l4 level1 lfo9;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;mso-ligatures:none;mso-fareast-language:PT-BR'>Quantidade de voucher(s) (patrono):<o:p></o:p></span></i></li>
                                </ul>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Cada voucher corresponde a uma passagem de ida e volta (exclusivamente sob a tarifa MAIS AZUL excluindo as classes A, B, E, F, G e Y) para qualquer trecho doméstico operado pela Azul. A validade para a realização da viagem de ida e volta é de 12 (doze) meses contados da data do protocolo do presente acordo nos autos judiciais e não será em nenhuma hipótese estendida e/ou renovada. A(s) reserva(s) está(ão) sujeita(s) à disponibilidade de assentos e deve(m) ser solicitada(s) com no mínimo 28 (vinte e oito) dias de antecedência da data pretendida para o voo de ida da viagem. O(s) referido(s) voucher(s) estará(ão) diretamente vinculado(s) à conta do Programa AZUL FIDELIDADE do(s) Autor(es) na(s) conta(s) indicada(s) acima.&#8221;<o:p></o:p></span></i></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><i><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'><o:p>&nbsp;</o:p></span></i></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Caso haja o aceite pedimos que nos envie os dados solicitados.<o:p></o:p></span></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Ficamos à disposição para discutir os detalhes desta proposta e para prestar qualquer esclarecimento adicional que se faça necessário.<o:p></o:p></span></p>
                                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;background:white'><span style='font-size:11.5pt;font-family:"Segoe UI",sans-serif;color:#242424;mso-ligatures:none;mso-fareast-language:PT-BR'>Atenciosamente,<o:p></o:p></span></p>
                                </div>
                                </body>
                        </html>


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
                                                    <td style="">
                                                        <a href="mailto:pm@nantesmello.com">
                                                            <img src="https://colosseo.com.br/email/nantes/davidson-galdino-infos.jpg" alt="David Galdino">
                                                        </a>
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                """
                        email.Send()
                        st.success(f'Email enviado com sucesso para {destinatario}.')
                    else:
                        st.write(f'Email enviado com sucesso para {destinatario}. (Simulação)')
                except Exception as e:
                    st.error(f'Erro ao enviar email para {destinatario}: {e}')
            else:
                st.warning('Por favor, verifique os destinatários')

    except Exception as e:
        st.error(f'Erro na preparação do envio dos emails: {e}')

# Upload do arquivo
df_file = st.file_uploader('Arraste aqui o relatório de acordo!', type=['csv', 'xlsx'])
if df_file is not None:
    df = pd.read_excel(df_file)
    st.dataframe(df.head(10))

    if st.button('Enviar'):
        enviar_email(df)
else:
    st.warning('Sem arquivo!')
