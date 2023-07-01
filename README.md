# User Registration

Portuguese version at the end

This is a Python program for user registration in an Excel file (.xlsx). The program uses the PySimpleGUI library to create a simple and user-friendly graphical interface.

## Prerequisites

Make sure you have the following libraries installed before running the program:

- PySimpleGUI
- openpyxl

You can install the libraries using pip:


## How to Use

1. Download or clone the repository.

2. Run the `cadastro.py` file to start the program.

3. The program window will open, allowing you to enter user information.

4. Fill in all the required fields: "Name" and "Age". Otherwise, an error message will be displayed.

5. Select the email providers that the user has by checking the corresponding checkboxes.

6. Choose whether the user accepts credit cards or not by selecting the appropriate option.

7. Click the "Submit" button to save the user's data.

8. The data will be added to an Excel file named "cadastro.xlsx" (if the file doesn't exist, it will be created automatically). The worksheet will contain the following columns: "Name", "Age", "Gmail", "Outlook", "Yahoo", "Accepts Card", and "Doesn't Accept Card".

9. The information will be centered in the worksheet for better visibility.

10. A message will be displayed informing that the data has been successfully saved.



# PORTUGUES

# Cadastro de Usuário

Este é um programa em Python para cadastro de usuários em um arquivo Excel (.xlsx). O programa utiliza a biblioteca PySimpleGUI para criar uma interface gráfica simples e amigável.

## Pré-requisitos

Certifique-se de ter as seguintes bibliotecas instaladas antes de executar o programa:

- PySimpleGUI
- openpyxl

Você pode instalar as bibliotecas usando o pip:

```
pip install PySimpleGUI
pip install openpyxl
```

## Como usar

1. Faça o download ou clone o repositório.

2. Execute o arquivo `cadastro.py` para iniciar o programa.

3. A janela do programa será aberta, permitindo que você insira informações do usuário.

4. Preencha todos os campos obrigatórios: "Nome" e "Idade". Caso contrário, uma mensagem de erro será exibida.

5. Selecione os provedores de e-mail que o usuário possui marcando as caixas de seleção correspondentes.

6. Escolha se o usuário aceita cartão ou não, selecionando a opção apropriada.

7. Clique no botão "Enviar" para salvar os dados do usuário.

8. Os dados serão adicionados a um arquivo Excel chamado "cadastro.xlsx" (caso o arquivo não exista, será criado automaticamente). A planilha conterá as seguintes colunas: "Nome", "Idade", "Gmail", "Outlook", "Yahoo", "Aceita Cartão" e "Não Aceita Cartão".

9. As informações serão alinhadas ao centro na planilha para melhor visualização.

10. Uma mensagem será exibida informando que os dados foram salvos com sucesso.

## Exemplo de Uso

```python
import PySimpleGUI as sg
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

# Código do programa omitido para maior clareza

tela = TelaPython()
tela.Iniciar()
```
