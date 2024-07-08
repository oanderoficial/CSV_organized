<h1>Organizar dados CSV </h1>

<p>Organizar dados de um arquivo CSV em uma planilha legível </p>

<strong> Importando bibliotecas </strong>
<p>Instalação:</p>

```
pip install pandas
```
```
pip install openpyxl
```
```
pip install tkinter
```
```python

import pandas as pd
import openpyxl
```



<strong> Leitura do CSV </strong>

```python 
arquivo = input("Digite o caminho do arquivo gerado pelo ServiceNow >>>")
dados = pd.read_csv(arquivo, encoding="latin-1")
```

<strong> Criação da pasta de trabalho e planilha </strong>

```python
job = openpyxl.Workbook()
planilha = job.active
```
<strong> Escrita dos cabeçalhos </strong>
```python
for i, colunas in enumerate(dados.columns):
    planilha.cell(row=1, column=i+1).value = colunas
```

<strong> Preenchimento dos dados </strong>
```python
for row_num, row in dados.iterrows():
    for col_num, colunas in enumerate(row):
        planilha.cell(row=row_num+2, column=col_num+1).value = colunas
```
<strong> Salvamento do arquivo Excel </strong> 
```python 
job.save('arquivo_organizado.xlsx')

print("Arquivo Excel organizado com sucesso!")
```
