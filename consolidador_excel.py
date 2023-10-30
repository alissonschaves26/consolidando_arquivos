import pandas as pd
import os
import datetime

data = datetime.datetime.now()
#Cria as colunas
colunas = [
    'Segmento',
    'País',
    'Produto',
    'Qtde de Unidades Vendidas',
    'Preço Unitário',
    'Valor Total',
    'Desconto',
    'Valor Total c/ Desconto',
    'Custo Total',
    'Lucro',
    'Data',
    'Mês',
    'Ano'
]

#Cria um datafame
df = pd.DataFrame(columns=colunas)

consolidado = pd.DataFrame(columns=colunas)

#Busca o nome dos arquivos
arquivos = os.listdir(r"E:\downloads\material_projeto01 (1)\planilhas")

#Realiza a consolidação dos arquivos
for excel in arquivos:
    
    if excel.endswith(".xlsx"):
        dados_arquivos = excel.split('-')
        segmento = dados_arquivos[0]
        pais = dados_arquivos[1].replace(".xlsx", "")
        
        try:        
            df = pd.read_excel(f"planilhas\\{excel}")
            df.insert(0, 'Segmento', segmento)
            df.insert(1, 'País', pais)
            consolidado = pd.concat([consolidado, df])
        except:
            with open("log_erro.txt", 'w') as arquivo:
                arquivo.write(f"Erro ao tentar consolidar o arquivo {excel}")
    else:
        with open("log_erro.txt", 'w') as arquivo:
            arquivo.write(f"O arquivo {excel} não é um arquivo excel valido")

#Exporta um dataframe para um arquivo excel
 
consolidado.to_excel(f"Report-consolidado-{data.strftime('%d-%m-%Y')}.xlsx", index=False, sheet_name='Report Consolidado')

consolidado.to_excel("Report.xlsx", index=False)
