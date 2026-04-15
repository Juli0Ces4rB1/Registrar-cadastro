# CRUD de Jogadores em Python

Projeto com interface grafica em `tkinter` para cadastrar, listar, atualizar e excluir jogadores em um arquivo Excel.
Os campos incluem nome, email, telefone, CPF, cidade, posicao, desempenho e foto.
A interface inclui busca por nome, CPF ou cidade, barra de status, layout em cards, formulario reorganizado em duas linhas, tabela ampliada com listras visuais, botao para selecionar foto, ordenacao por colunas, edicao por duplo clique, troca entre tema claro/escuro com salvamento da preferencia e melhor aproveitamento de tela.

## Requisitos

- Python 3.10 ou superior
- openpyxl
- tkinter

## Como executar

```bash
python atividade1.py
```

Na primeira execucao, o programa cria o arquivo `clientes.xlsx`.
O CPF recebe formatacao automatica no padrao `000.000.000-00`.
O telefone recebe formatacao automatica no padrao `(00) 00000-0000` ou `(00) 0000-0000`.
Voce pode trocar o logo padrao colocando um arquivo `logo.png` na pasta do projeto.
