import pandas as pd

dias = ['seg', 'ter', 'quar', 'quin', 'sex']

df = pd.read_excel('/home/einstein/Downloads/cenario01.xlsx', sheet_name="professor")
df.professor.fillna(method='ffill', inplace=True)

disc_df = pd.read_excel('input.ods', sheet_name="disciplinas")
disc_df['cod'] = disc_df['cod'].str.strip()
disc_df['nome'] = disc_df['nome'].str.strip()
disc_nome = disc_df.set_index("nome")

data = []
for k, v in df.iterrows():
    row = {}
    row['professor'] = v.professor
    row['início'] = v['horário'].split('-')[0] + ":00"
    row['fim'] = v['horário'].split('-')[1] + ":00"
    for dia in dias:
        row[dia] = '' if pd.isna(v[dia]) else 'X'
        if not pd.isna(v[dia]):
            row['nome_disciplina'] = v[dia]
            dn = v[dia].split('[')[0]
            row['cod_disciplina'] = disc_nome.loc[dn]['cod']
            row['ch'] = disc_nome.loc[dn]['ch']
            # print(dn)
    data.append(row)
    # print(row)

output = pd.DataFrame(data=data,columns=['cod_disciplina','nome_disciplina','ch','professor','seg', 'ter', 'quar', 'quin', 'sex','início','fim'])
with pd.ExcelWriter('output_conv.xlsx') as writer:
    output.to_excel(writer, sheet_name='t1')
