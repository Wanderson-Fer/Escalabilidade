import win32com.client as win32
import pandas as pd


def mail_reader():
    # ## Trecho de captura de cronograma diario
    # # Pré requisito: Tabela de cadastro de codigos

    # Criando referencia para a aplicação OutLook
    outlook = win32.Dispatch('outlook.application')
    # Iniciando Seção
    mapi = outlook.GetNameSpace("MAPI")

    # Capacitado apenas para uma conta
    # Verificando contas cadastradas e exibindo-as
    print('Contas cadastradas:')
    for account in mapi.Accounts:
        print(f'accounts - {account.DeliveryStore.DisplayName}')

    # Referência a pasta 'caixa de entrada' com os e-mails
    inbox = mapi.GetDefaultFolder(6)

    # Obtendo menssagens da pasta referênciada
    messages = inbox.Items

    # Indentificação do e-mail de programação por e-mail
    assunto = 'PACKING RN: Programação Diária'
    # Configurar filtro
    mailfiltro = f"@SQL=(urn:schemas:httpmail:subject LIKE '%{assunto}%')"
    # Aplicar filtro
    messages = messages.Restrict(mailfiltro)
    # Ordenando por data afim de capturar o mais recente
    messages.Sort('[ReceivedTime]', True)

    # Acessando tabela apartir do PRIMEIRO e-mail(mais recente)
    list_df_programacao = pd.read_html(messages[0].HTMLBody,
                                       match='Programação Produção de Pallets',
                                       decimal=',', header=4, )

    df_programacao = list_df_programacao[0]

    # Tabela da forma que está no e-mail
    return df_programacao


def get_programacao(df_cadastro_codigo, df_programacao=mail_reader()):
    """
    :param df_cadastro_codigo:
    :param df_programacao:
    :return df_programacao (consolidated):
    """

    col_programacao = {
        'Máquina': 'Descrição de linha',
        'Máquinas': 'Descrição de linha',
        'FERT': 'Código',
        'Fert': 'Código',
        'Configuração': 'Descrição',
        'Total': '1º Turno',
        'Total.1': '2º Turno',
        'Total.2': '3º Turno',
        '1º turno Total': '1º Turno',
        '2º turno Total': '2º Turno',
        '3º turno Total': '3º Turno',
    }

    # Renomeando colunas
    df_programacao = df_programacao.rename(columns=col_programacao)
    # Eliminando colunas
    df_programacao = df_programacao[list(set(col_programacao.values()))]

    # Tipagem
    df_programacao = df_programacao.astype({'Código': 'str'})

    # # Informações adicionais da planilha de código de MDO e Operação
    col_codigo = ['Código', 'MDO', 'Operação']
    df_programacao = pd.merge(df_programacao, df_cadastro_codigo[col_codigo],
                              on=['Código'], how='left')

    # # Códigos ("SKU's") sem 'Mão de Obra' cadastrada
    df_nok_cod = df_programacao.loc[df_programacao['MDO'].isnull()]

    df_programacao = df_programacao.loc[df_programacao['MDO'].notnull()]

    # ## 'df_programacao' consolidado.
    return df_programacao, df_nok_cod
