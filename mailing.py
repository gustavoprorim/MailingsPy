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

df['FONE1'] = df['FONE1'].apply(lambda x: re.sub(r'^\d{2}\s', '', str(x)))
df['FONE2'] = df['FONE2'].apply(lambda x: re.sub(r'^\d{2}\s', '', str(x)))

df.to_excel('padraomailing.xlsx', index=False)
