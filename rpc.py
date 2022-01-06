import pandas as pd


def corrigir(dados):
    dados = dados.replace('/', '').replace('.', '').replace('-', '')
    return dados

print('---------------------------------------------------------------------------------------------------------------')
print('          Conversor de arquivo Excel para TXT para importação dos dados no programa RPC da ANS\n')
print(' O Objetivo do programa é efetuar a conversão do arquivo utilizado para o preenchimento do RPC da ANS. \n'
      ' Com o arquivo transformado em txt será possível efetuar a importação, não será mais necessário efetuar o \n'
      'preenchimento manualmente, sendo assim, haverá um ganho tempo operacional importante. \n'
      ' - Para que o programa funcione corretamente o preenchimento dos dados deverão seguir um padrão, \n'
      'caso contrário poderá haver falha no processo, como informações faltantes ou incompletas. \n'
      ' - Será possível efetuar o download do layout do arquivo excel através do próprio programa.\n'
      ' - O tipo de registro C4 está desabilitado, pois não é efetuado esse preenchimento hoje no preenchimento manual.'
      ' \n'
      'As informações usadas para efetuar a validação do arquivo foram retiradas do manual do RPC disponibilizado'
      ' pela ANS, versão 3.1.4 de 31/08/2016, disponível no site: \n'
      'https://www.gov.br/ans/pt-br/centrais-de-conteudo/manuais-do-portal-operadoras/reajuste-de-planos-coletivos-rpc')
print('----------------------------------------------------------------------------------------------------------------'
      ' \n \n \n')

digite = ''

while digite == '':
    digite = input('Digite 1 para continuar o programa de conversão,\n'
                   'Digite 2 para gerar o arquivo com o layout do excel\n'
                   'Digite 3 para sair do programa: \n')
    if digite == '1':
        # Dados para montar o registro C1
        rg_ans = input('Informar o registro da Operadora na ANS: ')
        cnpj = input('Informar o CNPJ da Operadora: ')
        razao = input('Informar a Razão Social da Operadora: ')
        dt_comunicado = input('Informar a data do Cadastramento do Comunicado (dd/mm/aaaa): ')

        # Informar o arquivo
        arq = input('Informar o arquivo excel com os dados: ')
        arq = arq.replace('\\', '/')

        # Informar onde salvar o arquivo
        saida = input('Informar o diretório para salvar o arquivo txt: ')
        saida = saida.replace('\\', '/')

        print('-----------------------------------------------------------')
        print('Processando...')

        # Criar o dataframe
        df_temp = pd.read_excel(f'{arq}')
        df = df_temp.copy()

        # Alterando os campos com data
        cp_data = ['Data Base', 'Início Aplicação', 'Fim Aplicação', 'Início Análise', 'Fim Análise']

        for x in cp_data:
            df[f'{x}'] = df[f'{x}'].dt.strftime('%m/%Y')
        for x in cp_data:
            df[f'{x}'] = df[f'{x}'].astype('str').apply(corrigir)

        # Criando a coluna com o sinal do Percentual de Reajuste Informando
        for x in range(0, len(df)):
            if df.loc[x, 'Percentual'] >= 0:
                df.loc[x, 'Sinal Percentual Reajuste'] = 'P'
            else:
                df.loc[x, 'Sinal Percentual Reajuste'] = 'N'

        # Transformando a coluna reajuste em str e retirando os caracteres desnecessários
        df['Percentual'] = df['Percentual'].astype('str')
        df['Percentual'] = df['Percentual'].apply(corrigir)

        df['Cod. Contratante'] = df['Cod. Contratante'].astype('str')

        # Complentando o valor na coluna Percentual
        for x in range(0, len(df)):
            if len(df.loc[x, 'Percentual']) == 2:
                df.loc[x, 'Percentual'] = df.loc[x, 'Percentual'] + '0'
            elif len(df.loc[x, 'Percentual']) == 1:
                df.loc[x, 'Percentual'] = df.loc[x, 'Percentual'] + '00'

        # Padronizar colunas com letra maiuscula
        cl_upper = ['Reajuste Linear', 'Introdução de Franquia / Co-participação']

        for x in cl_upper:
            df[f'{x}'] = df[f'{x}'].str.upper()

        # Encontrando o dado Dispersão atraves da coluna Abrangência
        for x in range(0, len(df)):
            if df.loc[x, 'Abrangência'] == 'NACIONAL':
                df.loc[x, 'Dispersão'] = 'M'
            else:
                df.loc[x, 'Dispersão'] = 'E'

        # Adcionando o espaçamento no campo
        df['Descrição Plano'] = df['Descrição Plano'].apply(lambda x: x.ljust(60))

        # Criando a linha C1
        c1_temp = ['C1', '{:0>6}'.format(rg_ans), '{:0>14}'.format(cnpj), '{:<60}'.format(razao),
                   dt_comunicado.replace('/', '')]
        c1 = ''.join(c1_temp)

        # As informações da linha C4 não são enviadas
        c4 = 'C4000000000000000000000000000000000000000000000'

        # Criando as linhas C2 e C3
        c2_temp = []
        c3_temp = []
        arquivo = []

        arquivo.append(c1)
        for x in range(0, len(df)):
            c2_temp = ['C2', '{:0>9}'.format(df.loc[x, 'Registro Plano']), df.loc[x, 'Descrição Plano'],
                       '{:>20}'.format(df.loc[x, 'Cod. Contratante']),
                       df.loc[x, 'Modalidade de Contratação'].astype('str'), df.loc[x, 'Início Aplicação'],
                       df.loc[x, 'Fim Aplicação'], df.loc[x, 'Início Análise'], df.loc[x, 'Fim Análise'],
                       '{:0>10}'.format(df.loc[x, 'Benef. Ativos']), df.loc[x, 'Sinal Percentual Reajuste'],
                       '{:0>6}'.format(df.loc[x, 'Percentual']), 'SP', df.loc[x, 'Dispersão'],
                       df.loc[x, 'Característica do Reajuste'].astype('str'),
                       df.loc[x, 'Reajuste Linear'], df.loc[x, 'Introdução de Franquia / Co-participação'], 'P',
                       '000000',
                       df.loc[x, 'Código do Plano'].astype('str')]
            temp = ''.join(c2_temp)
            c3_temp = ['C3', '001', df.loc[x, 'Justificativa Técnica']]
            temp_3 = ''.join(c3_temp)
            arquivo.append(temp)
            arquivo.append(temp_3)
            #arquivo.append(c4)

        # Criando o registro C9
        '''c9 = ['C9', '{:0>8}'.format(len(arquivo) + 1), '{:0>6}'.format(len(df)), '{:0>6}'.format(len(df)),
              '{:0>7}'.format(len(df))]'''
        c9 = ['C9', '{:0>8}'.format(len(arquivo) + 1), '{:0>6}'.format(len(df)), '{:0>6}'.format(len(df)),
              '0000000']
        c9_temp = ''.join(c9)
        arquivo.append(c9_temp)

        # para complentar o nome do arquivo de saída
        nome = dt_comunicado.replace('/', '')

        # Gerando o arquivo txt
        txt = open(f'{saida}/{nome}_RPC.txt', 'w')

        for x in arquivo:
            txt.write(f'{x}\n')
        txt.close()

    elif digite == '2':
        coluna = ['Data Base', 'Início Aplicação', 'Fim Aplicação', 'Início Análise', 'Fim Análise', 'Cod. Contratante',
                  'Nome Contratante', 'Modalidade', 'Termo', 'Percentual', 'Comp Grp Contratante', 'Benef. Ativos',
                  'Registro Plano',
                  'Descrição Plano', 'Abrangência', 'Justificativa Técnica', 'Modalidade de Contratação',
                  'Característica do Reajuste', 'Reajuste Linear', 'Introdução de Franquia / Co-participação',
                  'Percentual de reajuste Franq', 'Código do Plano', ]
        dir_layout = input('Digite o diretório para salvar o arquivo com o layout excel: ')
        dir_layout = dir_layout.replace('\\', '/')
        layout = pd.DataFrame(columns=coluna)
        layout.to_excel(f'{dir_layout}/layout.xlsx', index=False)
    elif digite == '3':
        break
    else:
        digite = ''
