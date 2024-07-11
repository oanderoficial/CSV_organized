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
from tkinter import *
import tkinter.messagebox as messagebox
from tkinter import ttk
from tkinter import filedialog
```



<strong> Leitura do CSV </strong>

```python 
  def run (self):
    # Leitura do CSV
        #arquivo = input("Digite o caminho do arquivo gerado pelo ServiceNow >>>")
        try:
            file_path = filedialog.askopenfilename(title="Digite o caminho do arquivo gerado pelo ServiceNow >>>", filetypes=[("csv", "*.csv")])
            dados = pd.read_csv(file_path, encoding="latin-1")
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
            messagebox.showinfo('Sucesso, dados organizados com sucesso!')
        except:
            messagebox.showerror("Erro", f"Ocorreu um erro ao carregar o arquivo csv:")
             
```

```python
if __name__ == "__main__":
    run = MainExcel()
    run.run()
```
