PLANILHA_TRILHAS = 'dados/trilhas_formativas.xlsx'
PLANILHA_PRESENCA = 'dados/lista_presenca_trilhas_formativas.xlsx'

COLUNAS_ALUNOS = {
    'aluno': 'Aluno',
    'matricula': 'Matricula',
    'cpf': 'CPF',
    'mae': 'Mae',
    'pai': 'Pai',
    'turma': 'Turma',
    'telefone': 'Telefone',
    'direcao': 'Direcao'
}
CAMPOS_BUSCA_DADOS = ["Aluno", "Matricula", "CPF", "Mae", "Pai", "Turma", "Telefone", "Escola"]


# Configurações da Interface
COR_AZUL_ESCURO = '#1976D2'
COR_BRANCA = '#FFFFFF'
COR_CINZA_CLARO = '#ECEFF1'

FONTE_TITULO = ('Helvetica', 14, 'bold')
FONTE_PRINCIPAL = ('Helvetica', 10)

JANELA_LARGURA = 800
JANELA_ALTURA = 600
IMAGEM_FUNDO = 'imagens_menu/fundo_menu.png'

# Mensagens de Erro
ERRO_ARQUIVO_NAO_ENCONTRADO = "Erro: Um ou mais arquivos de planilha não foram encontrados. Verifique os caminhos 'dados/'."
ERRO_DADOS = "Erro no processamento dos dados. Verifique o formato das planilhas."