# 📊 Projeto: Análise de Trilhas Formativas

## 💻 Descrição do Projeto
Este é um sistema de desktop simples (utilizando Tkinter e Pandas) desenvolvido em Python para facilitar a análise de dados cadastrais e o cálculo da frequência de alunos em diversas Oficinas de Trilhas Formativas.

O sistema processa planilhas de dados (alunos e listas de presença) e oferece duas funcionalidades principais:
1.  **Busca Cadastral:** Pesquisa por Aluno, Matrícula ou Escola e exibe dados completos.
2.  **Cálculo de Frequência:** Calcula o percentual de presença por aluno/oficina e lista a presença completa por dia.

## ✨ Funcionalidades

* **Busca Flexível:** Pesquisa de alunos por diversos campos (Aluno, Matrícula, CPF, Escola, etc.).
* **Normalização de Busca:** A busca é insensível a acentos e letras maiúsculas/minúsculas.
* **Cálculo Preciso:** Calcula a frequência percentual de cada aluno em cada oficina (corrigindo problemas de contagem dupla).
* **Dados Detalhados:** Exibe a lista completa de alunos presentes em cada dia da oficina.

## Instalação de Dependências
Abra o terminal na pasta raiz do projeto e execute:
pip install pandas openpyxl Pillow

## 🛠️ Estrutura de Arquivos

ProjetoTrilhasFormativas/
├── dados/
│   ├── trilhas_formativas.xlsx   # (Dados cadastrais)
│   └── lista_presenca_trilhas_formativas.xlsx # (Dados de presença)
├── imagens_menu/
│   └── fundo_menu.png            # (Imagem de fundo da tela inicial)
├── Const.py                      # (Arquivo de constantes e configurações)
└── main_app.py                   # (Código principal da aplicação)
