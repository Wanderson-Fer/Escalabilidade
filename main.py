import getpass
import sqlite3
import pandas as pd
import datetime as dt

# Captura de Cronograma por e-mail
import win32com.client as win32

from frequencia import get_frequencia

# obtendo o nome do usuário atual
user = getpass.getuser()
# Definindo o diretório para as planilhas
path = f'C:\\Users\\{user}\\Procter and Gamble\\Grupo Check List - Bases de dados\\'

# Definindo acessos para a planilha de cadastros
cadastro_file = 'planilhas\\df_Cadastro.xlsx'

# Tabela com o cadastro de habilidade por nome de funcionário
df_cadastro_habilidade = pd.read_excel(path + cadastro_file, sheet_name='dHabilidade', dtype=str)
# Tabela com o cadastro de códigos produtos e materia-prima
df_cadastro_codigo = pd.read_excel(path + cadastro_file, sheet_name='dCódigo', dtype=str)

# Definincdo acesso para tabela de feriados nacionais
feriados_file = 'planilhas\\df_Feriado.xlsx'

# Tabela com os fériados nacionais relevantes
df_feriados = pd.read_excel(path + feriados_file)

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

# Referência as pastas 'caixa de entrada' com os e-mails
inbox = mapi.GetDefaultFolder(6)

# Obtendo menssagens da pasta referênciada
messages = inbox.Items

# Indentificação do e-mail de programação por e-mail
assunto = 'PACKING RN: Programação Diária'
# Configurar e aplicar filtro
mailfiltro = f"@SQL=(urn:schemas:httpmail:subject LIKE '%{assunto}%')"
messages = messages.Restrict(mailfiltro)
# Ordenando por data afim de capturar o mais recente
messages.Sort('[ReceivedTime]', True)

# Acessando tabela apartir do primeiro e-mail(mais recente)
list_df_programacao = pd.read_html(messages[0].HTMLBody,
                                   match='Programação Produção de Pallets',
                                   decimal=',', header=4, )

df_programacao = list_df_programacao[0]

# ## Tratamento da tabela

col_programacao = {
    'Máquina': 'Descrição de linha',
    'FERT': 'Código',
    'Configuração': 'Descrição',
    'Total': 'Turno 1',
    'Total.1': 'Turno 2',
    'Total.2': 'Turno 3'
}

# Renomeando colunas
df_programacao = df_programacao.rename(columns=col_programacao)
# Eliminando colunas
df_programacao = df_programacao[list(col_programacao.values())]

# Tipagem
df_programacao = df_programacao.astype({'Código': 'str'})

# # Informações adicionais
col_codigo = ['Código', 'MDO', 'Posição', 'Operação']
df_programacao = pd.merge(df_programacao, df_cadastro_codigo[col_codigo],
                          on=['Código'], how='left')

# # Códigos ("SKU's") sem 'Mão de Obra' cadastrada
df_nok_cod = df_programacao.loc[df_programacao['MDO'].isnull()]

df_programacao = df_programacao.loc[df_programacao['MDO'].notnull()]

# ## 'df_programacao' consolidado.

df_programacao['Prioridade'] = ''
df_programacao.loc[
    df_programacao['Operação'] == 'SHERINKADORA',
    'Prioridade'
] = 'X'

df_programacao.sort_values(by=['Prioridade', 'Operação', 'Posição'],
                           ascending=False, inplace=True)

list_turno = [
    'Turno 1',
    'Turno 2',
    'Turno 3'
]

# Compatibilidade de colunas
df_cadastro_habilidade = df_cadastro_habilidade.rename(columns={'Código': 'Código Func'})

# 'df_frequencia' para lógica de frequencia
df_frequencia = get_frequencia()

df_prog_today = pd.DataFrame(columns=['Descrição de linha', 'Código', 'TURNO', 'Código Func', 'Data'])

for turno in list_turno:
    print(turno)
    df_temp_fturno = df_cadastro_habilidade.loc[df_cadastro_habilidade['TURNO'] == turno]
    df_temp_pturno = df_programacao.loc[df_programacao[turno].notnull()]
    list_geral_func = []

    for indice in df_temp_pturno.index:
        desc_linha, cod_sku = df_temp_pturno.loc[indice, ['Descrição de linha', 'Código']].values

        df_temp_ffreq = df_temp_fturno.merge(
            df_frequencia.loc[
                (df_frequencia['Descrição de linha'] == desc_linha) &
                (df_frequencia['Código'] == cod_sku),
                ['Código Func', 'Frequencia', 'Data']
            ],
            on=['Código Func'],
            how='left'
        )

        df_temp_ffreq.sort_values(['Data', 'Frequencia'], inplace=True, na_position='first')

        operacao, mdo = df_temp_pturno.loc[indice, ['Operação', 'MDO']]
        if df_temp_pturno.loc[indice, 'Prioridade'] == 'X':
            list_func = list(df_temp_ffreq.loc[df_temp_ffreq[operacao] == 'X', 'Código Func'])
        else:
            list_func = list(df_temp_ffreq['Código Func'])

        repetidos = [func for func in list_func if func in list_geral_func]
        if repetidos:
            for func in repetidos:
                list_func.remove(func)

        # Lógica de frequencia

        if int(mdo) <= len(list_func):
            list_func = list_func[0:int(mdo)]

        list_geral_func += list_func

        for cod_func in list_func:
            df_prog_today.loc[len(df_prog_today)] = [
                    desc_linha,
                    cod_sku,
                    turno,
                    cod_func,
                    dt.date(day=24, month=1, year=2023).strftime('%d/%m/%Y')
                    # dt.date.today().strftime('%d/%m/%Y')
            ]

# Definindo acessi para DB de programação de aeroporto
db_airport_file = 'bd_sqlite\\db_airport.db'
conn = sqlite3.connect(path + db_airport_file)

# Realizando gravação no historico de programação do aeroporto
df_prog_today.to_sql('hist_programacao', conn, if_exists='append', index=False)

# Tabela de histórivo de programaçao
df_programacao_historico = pd.read_sql('SELECT * FROM hist_programacao', conn)

# Fechando o conector
conn.close()
