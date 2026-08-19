import pandas as pd

df_python = pd.read_csv(r'resultados\Indescola_Escolas.csv', sep=';')

print('Calculando quantidade de escolas sem escala...')
total = df_python['Nivel_Infraestrutura'].value_counts().sum()
sem_escala = df_python['Nivel_Infraestrutura'].value_counts().get('Sem Escala',0)
porcentagem = sem_escala/total


print(f'Numero de escolas sem escala/total de escolas: {sem_escala}/{total} ({porcentagem:.2%})')
