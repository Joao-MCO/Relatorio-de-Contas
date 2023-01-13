import pandas as pd
import numpy as np
import ArqFinder
import os
import time
import datetime
import matplotlib.pyplot as plt
import seaborn as sns
import win32com.client as client
import xlsxwriter

# Salva o tempo incial para calcular o tempo gasto
t = time.time()

# Localizar Aqruivo
path = ArqFinder.open()

# Importar Arquivo e limpa colunas com informação desnecessária
tabela_imp = pd.read_excel(path)
tabela = pd.read_excel(path)
tabela = tabela.dropna(axis=1, how="all")
tabela = tabela.drop(["HORA","LOGIN","DOCPG","NFS","DESCRESPECIE","STATUS"],axis=1)

# Calcular o SubTotal de cada dia
grp_tabela = tabela[["DATA","VALORCRED","VALORDEB"]].groupby("DATA", as_index=False).sum()
grp_tabela['VARIACAO'] = grp_tabela['VALORCRED'] - grp_tabela['VALORDEB']

# Transformação em uma única tabela
dataRef = tabela["DATA"].loc[1] # Usa uma data para saber quando termina o dia
final = pd.DataFrame()
final = final.append(tabela.loc[0], ignore_index=True) # Insere o primeiro registro que é o saldo do último dia do mês passado
final["VARIACAO"] = final["VALORCRED"] # Insere a coluna VARIACAO
grp_tabela["SALDO"] = tabela["SALDO"].loc[0]
aux = 0 # Variável como índice usada para inserir o dado do DataFrame: grp_tabela
for i in range(1, tabela.shape[0]+grp_tabela.shape[0]-1):
    # [i-aux] é o indice usado para inserir o dado do DataFrame: tabela. Enquanto a condição for satisfeita, há dados no DataFrame tabela
    if(i-aux<tabela.shape[0]): 
        # Caso a date de  [i - aux] é do mesmo dia da referência, insere normal do DataFrame: tabela
        if(dataRef == tabela["DATA"].loc[i-aux]):
            final = final.append(tabela.loc[i-aux], ignore_index=True)
        # Caso [i - aux] é de um dia diferente do dia da referência, insere do DataFrame: grp_tabela
        else:
            final = final.append(grp_tabela.loc[aux+1], ignore_index=True)
            final["SALDO"].loc[i] = final["SALDO"].loc[i-1] # Mantém o último saldo
            dataRef = tabela["DATA"].loc[i-aux] # Atualiza a referência para o dia de [i - aux]
            dia = final["DATA"].loc[i] # Variável usada na descrição do SubTotal
            final["DESCRICAO"].loc[i] = "SALDO FINAL DO DIA {}/{}/{}".format(dia.day, dia.month, dia.year) # Descrição do SubTotal
            grp_tabela["SALDO"].loc[aux+1] = final["SALDO"].loc[i] # Joga o Saldo para o GroupBy para ser analisado
            aux += 1  # Atualiza o aux
    # No caso da condição não ser satisfeita, significa o fim dos dados do DataFrame tabela. Insere o SubTotal do último dia
    else:
        final = final.append(grp_tabela.loc[aux+1], ignore_index=True)
        final["SALDO"].loc[i] = final["SALDO"].loc[i-1] # Mantém o último saldo
        dia = final["DATA"].loc[i] # Variável usada na descrição do SubTotal
        final["DESCRICAO"].loc[i] = "SALDO FINAL DO DIA {}/{}/{}".format(dia.day, dia.month, dia.year) # Descrição do SubTotal
        grp_tabela["SALDO"].loc[aux+1] = final["SALDO"].loc[i] # Joga o Saldo para o GroupBy para ser analisado
        
# Recebe o dia do inicio e o dia do fim da consulta
inicio = final["DATA"].loc[1]
fim = final["DATA"].loc[final.shape[0]-1]
bruto = grp_tabela["DATA"].loc[grp_tabela["VALORCRED"].idxmax()]
liquido = final["DATA"].loc[final["VARIACAO"].idxmax()] 
totalLiquido = final["VARIACAO"].max()
maxMov = tabela["DATA"].value_counts().idxmax()

# Cria um dicionário inserindo uma linha vazia e o saldo final     
x = {
    "DATA": [np.nan, np.nan],
    "VALORCRED": [np.nan,tabela["VALORCRED"].sum()],
    "VALORDEB": [np.nan, tabela["VALORDEB"].sum()],
    "SALDO": [np.nan,tabela["SALDO"].loc[tabela.shape[0]-1]],
    "DESCRICAO": [np.nan, "SALDO FINAL DO PERIODO: " + "{}/{}/{} - {}/{}/{}".format(inicio.day, inicio.month, inicio.year,
    fim.day, fim.month, fim.year)],
    "VARIACAO": [np.nan, tabela["VALORCRED"].sum() - tabela["VALORDEB"].sum()]
}
# Transforma o dicionario em um DataFrame
total = pd.DataFrame(x)

# Concatena o DataFrame
final = final.append(total, ignore_index=True)

# Ordena as Colunas
variacao = final.pop("VARIACAO")
final.insert(3, "VARIACAO", variacao)

# Convertendo as Colunas para o tipo correto
final["DATA"] = pd.to_datetime(final["DATA"], errors="coerce")
final["VARIACAO"] = pd.to_numeric(final["VARIACAO"], errors="coerce")

# Criação da Pasta Alterados DD-MM-AAAA
pathAux = path.split(sep='/') # Separo o diretório em cada barra para poder criar uma pasta nova
path = ''

for i in range(len(pathAux)-1):
    path = path + pathAux[i] + '/' # Volta as pastas anteriores para o caminho

path = path + 'Alterado {}-{}-{}'.format(datetime.date.today().day,datetime.date.today().month, datetime.date.today().year) # Adiciona a pasta ao caminho

# Tenta criar a pasta. Se ela existe, a mensagem: Ja existe a pasta é impressa
try:
    os.mkdir(path=path)
except:
    print("Ja existe a pasta")

# Sumario
sum = {
    "DESCRICAO": ["Dia com maior receita bruta","Dia com maior receita liquida","Dia com mais movimentacoes",
    "Media Diaria da Variacao", "Mediana da Variacao", "Saldo Inicial", "Saldo Final", "Variacao Total",
    "Total de Movimentações", "Media de Movimentações"],

    "VALOR": [grp_tabela["VALORCRED"].max(), totalLiquido, tabela["DATA"].value_counts().max(), final["VARIACAO"].mean(), 
    final["VARIACAO"].median(), final["SALDO"].loc[0], final["SALDO"].loc[final.shape[0]-1], 
    final["VARIACAO"].loc[final.shape[0]-1], tabela["DATA"].value_counts().sum(), tabela["DATA"].value_counts().mean()],

    "DIA": [bruto, liquido, maxMov, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan]
}
sumario = pd.DataFrame(sum)

# Gráficos
sns.set_theme(style="ticks")
xlim = [0, len(grp_tabela)]

plt.figure("CRED X DEB", figsize=(12,12))
if(grp_tabela["VALORDEB"].max() > grp_tabela["VALORCRED"].max()):
    plt.yticks(np.arange(0, grp_tabela["VALORDEB"].max(), (grp_tabela["VALORDEB"].max())//10))
else:
    plt.yticks(np.arange(0, grp_tabela["VALORCRED"].max(), (grp_tabela["VALORCRED"].max())//10))
plt.xticks(rotation=45)
plt.title("CRED X DEB")
sns.lineplot(data=grp_tabela, x="DATA", y="VALORCRED", legend=True)
sns.lineplot(data=grp_tabela, x="DATA", y="VALORDEB", legend=True)
plt.savefig('CRED X DEB.png')
cxd = pd.DataFrame()

plt.figure("RENDA LIQUIDA", figsize=(12,12))    
plt.yticks(np.arange(grp_tabela["VARIACAO"].min() - 10000, grp_tabela["VARIACAO"].max() + 10000, 
    (grp_tabela["VARIACAO"].max() - grp_tabela["VARIACAO"].min())//10))
plt.xticks(rotation=45)
plt.title("RENDA LIQUIDA")
sns.lineplot(data=grp_tabela, x="DATA", y="VARIACAO")
plt.savefig('VARIACAO.png')
var = pd.DataFrame()

plt.figure("SALDO", figsize=(12,12))
plt.yticks(np.arange(grp_tabela["SALDO"].min() - 10000, grp_tabela["SALDO"].max() + 10000, 
    (grp_tabela["SALDO"].max() - grp_tabela["SALDO"].min())//10))
plt.xticks(rotation=45)
plt.title("SALDO")
sns.lineplot(data=grp_tabela, x="DATA", y="SALDO")
plt.savefig('SALDO.png')
sal = pd.DataFrame()

final.style.format({"VALORCRED":'R$ {:,.2f}',"VALORDEB":'R$ {:,.2f}',"SALDO":'R$ {:,.2f}',"VARIACAO":'R$ {:,.2f}'})
grp_tabela.style.format({"VALORCRED":'R$ {:,.2f}',"VALORDEB":'R$ {:,.2f}',"SALDO":'R$ {:,.2f}',"VARIACAO":'R$ {:,.2f}'})

# Exportação para Excel
path = path + '/' + pathAux[len(pathAux)-1] # Adiciona o arquivo com o mesmo nome do importado ao caminho
try:
    writer = pd.ExcelWriter(path, engine = 'xlsxwriter')
    sumario.to_excel(writer, "Sumario dos Dados")
    final.to_excel(writer, "Relatorio com SubTotal")
    tabela_imp.to_excel(writer, "Relatorio exportado do IC")
    cxd.to_excel(writer, "CRED X DEB")
    ws = writer.sheets["CRED X DEB"]
    ws.insert_image('C3','CRED X DEB.png')
    var.to_excel(writer, "VARIACAO")
    ws = writer.sheets["VARIACAO"]
    ws.insert_image('C3','VARIACAO.png')
    sal.to_excel(writer, "SALDO")
    ws = writer.sheets["SALDO"]
    ws.insert_image('C3','SALDO.png')
    writer.close()
    print("Arquivo Salvo")
except:
    print("Não foi possível salvar o arquivo corretamente!")