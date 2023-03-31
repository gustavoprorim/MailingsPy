import pandas as pd
import sys
import re

if len(sys.argv) < 2:
    print("Por favor, forneça o caminho para o arquivo como argumento.")
    sys.exit()

filename = sys.argv[1]

df = pd.read_excel(filename)

col_order = ['NOME', 'NR_CPF', 'DS_ENDERECO', 'NR_ENDERECO', 'DS_COMPLEMENTO',
             'CD_CEP', 'DS_BAIRRO', 'MUNICIPIO', 'ESTADO', 'DDD1', 'FONE1', 'DDD2', 'FONE2', 'DT_NASC', 'DS_EMAIL']
df = df[col_order]

df['NR_CPF'] = df['NR_CPF'].astype(str)
df['NR_CPF'] = df['NR_CPF'].apply(lambda x: x.zfill(11))
df['NR_CPF'] = df['NR_CPF'].apply(lambda x: re.sub(r'(\d{3})(\d{3})(\d{3})(\d{2})', r'\1.\2.\3-\4', x))

df['DDD'] = df[['DDD1', 'DDD2']].apply(lambda x: x.dropna().astype(str).str.replace('-', '').str.zfill(2).str.extract(r'(\d{2})', expand=False).iloc[0] if len(x.dropna().astype(str).str.replace('-', '').str.zfill(2).str.extract(r'(\d{2})', expand=False)) > 0 else None, axis=1)

df['FONE'] = df[['FONE1', 'FONE2']].apply(lambda x: x.dropna().astype(str).str.replace(r'^\d{2}\s', '').str.extract(r'(\d{8,9})', expand=False).iloc[0] if len(x.dropna().astype(str).str.replace(r'^\d{2}\s', '').str.extract(r'(\d{8,9})', expand=False)) > 0 else None, axis=1)

df = df.drop(['DDD1', 'DDD2', 'FONE1', 'FONE2'], axis=1)

df = df.drop_duplicates(subset=['NOME', 'NR_CPF'])

df = df.dropna(subset=['DDD', 'FONE'])

df.to_excel('padraomailing.xlsx', index=False)
