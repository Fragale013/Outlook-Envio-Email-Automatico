#Primeiramente vamos importar as bibliotecas necessárias
#Para utilizar esse código, você precisa ter uma versão do Microsoft Outlook devidamente configurado na sua máquina Windows

import win32com.client as win32
import pandas as pd
from time import sleep
import xlrd

#Vamos inicialmente carregar a tabela Excel no nosso código, no meu caso ela se encontra nesse diretório
df = pd.read_excel('C:/Users/Usuario/OneDrive - Minha Empresa/Projeto/Envio_email/Disparo.xlsb')

#Dentro da minha tabela, os serviços que tiverem com o campo StatusFinal = 'PENDENTE' deverão ter email enviado para o responsável da área
quant_pend = df[df['StatusFinal'] == 'PENDENTE']['StatusFinal'].count()

#Agora vem o mecanismo de decisão, preciso no meu caso enviar para um mailing específico os serviços da cidade do Rio de Janeiro:
if quant_pend > 0:
    df_pend = df[(df['StatusFinal'] == 'PENDENTE')] #Aqui estamos filtrando a tabela somente com os casos pendentes
    df_pend.fillna(value="", inplace=True) #Nesse ponto estamos tratando os possíveis dados nulos
    #Agora vamos percorrer as linhas:
    for index, linha in df_pend.iterrows():
        if linha['Cidade'] == 'Rio de Janeiro':
            dt_agenda_xl = int(linha['DT_AGENDA'])
            dt_agenda = xlrd.xldate_as_datetime(dt_agenda_xl, 0)
            dt_agenda_convert = dt_agenda.strftime('%d/%m/%Y')
            #As 3 linhas acima foram necessárias para tratar os dados provenientes de data em excel para o datetime conhecido data/mês/ano
            
            #Agora iremos escrever a mensagem, observe como o mailing é fixo, porém a mensagem será alterada conforme a linha da tabela que for lida
            #O mailing também pode ser variável, só deixar uma linha na tabela com o mailing necessário em cada caso e chamar ele aqui como variável como os demais
            mail = outlook.CreateItem(0)
            mail.To = 'fulano@empresaqualquer.com.br; ciclano@empresaqualquer.com.br; deutrano@empresaqualquer.com.br'
            mail.CC = 'usuarioqualquer@empresa2.com.br; usuarioqualquer2@empresa2.com.br'
            mail.Subject = f'Centro de Operações | ESTADO {linha["ESTADO"]} | {linha["CIDADE"]} {linha["AREA_DESCRICAO"]} | Oportunidade de acesso ao imóvel {linha["Código GED"]}'
            mail.Body = f'''
                Caro gestor,
                Identificamos oportunidade de acesso ao prédio na {linha['END_COMPLETO']} {linha['ID_COMPL1']} {linha['COMPL1_DESCR']}, através da VISITA TÉCNICA abaixo:
                Contrato: {linha['NUM_CONTRATO']}
                Data Agenda: {dt_agenda_convert}
                Horário da Agenda: {linha['AGENDA_DESCR']}
                Sugerimos que seja enviado um técnico de atendimento junto ao endereço para realizar o serviço pendente do ticket {linha['Ticket']}.
                Atenciosamente,
                Centro operacional - Atendimento e apoio a campo
                Planejamento, automação e inteligência de dados.
                '''
            mail.Send()
            print('Email enviado com sucesso')
        else:
            pass
    sleep(10)
else:
    print("Nenhuma oportunidade pendente")
    sleep(10)
