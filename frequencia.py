import sqlite3
import pandas as pd
import datetime as dt
import getpass

def get_frequencia():
    # obtendo o nome do usuário atual
    user = getpass.getuser()
    # Definindo o diretório para as planilhas
    path = f'C:\\Users\\{user}\\Procter and Gamble\\Grupo Check List - Bases de dados\\'

    # Definindo acessi para DB de programação de aeroporto
    db_airport_file = 'bd_sqlite\\db_airport.db'
    conn = sqlite3.connect(path + db_airport_file)

    # Tabela de histórivo de programaçao
    df_programacao_historico = pd.read_sql('SELECT * FROM hist_programacao', conn)

    # Fechando o conector
    conn.close()

    df_programacao_historico['Data'] = pd.to_datetime(
        df_programacao_historico['Data'],
        dayfirst=True
    ).dt.date

    df_programacao_mensal = df_programacao_historico.loc[
        df_programacao_historico['Data'] >
        (dt.date.today() - dt.timedelta(30))
    ]

    df_frequencia = df_programacao_mensal[
        ['Descrição de linha', 'Código', 'Código Func', 'Data']
    ].groupby(
        ['Descrição de linha', 'Código', 'Código Func']
    ).count().reset_index()

    df_frequencia = df_frequencia.rename(columns={'Data': 'Frequencia'})

    df_frequencia = df_frequencia.merge(
        df_programacao_mensal[['Descrição de linha', 'Código', 'Código Func', 'Data']],
        on=['Descrição de linha', 'Código', 'Código Func'],
        how='left'
    )

    df_frequencia.sort_values(['Descrição de linha', 'Data'], ascending=False, inplace=True)

    df_frequencia.drop_duplicates(['Descrição de linha', 'Código', 'Código Func'], inplace=True)

    return df_frequencia
