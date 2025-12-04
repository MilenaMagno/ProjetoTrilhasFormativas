import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import unicodedata
import numpy as np

try:
    from Const import *
except ImportError:
    messagebox.showerror("Erro Crítico",
                         "O arquivo 'Const.py' não foi encontrado. Certifique-se de que ele está na mesma pasta.")
    exit()
except Exception as e:
    messagebox.showerror("Erro de Importação", f"Erro desconhecido ao carregar Const.py: {e}")
    exit()

try:
    from PIL import Image, ImageTk
except ImportError:
    messagebox.showerror("Erro de Dependência",
                         "A biblioteca 'Pillow (PIL)' não está instalada. Execute 'pip install Pillow' no seu terminal.")
    exit()


# FUNÇÕES AUXILIARES

def normalize_text(text):
    """
    Remove acentos, converte para minúsculas, remove espaços extras e trata NaN/None.
    É crucial para a busca de alunos e escolas.
    """
    if pd.isna(text) or text is None:
        return ""

    try:
        text_str = str(text).strip()
    except Exception:
        return ""

    normalized = unicodedata.normalize('NFD', text_str).encode('ascii', 'ignore').decode("utf-8")
    normalized = ' '.join(normalized.split())
    return normalized.lower()


def create_menu_button(parent, text, command):
    """Cria um botão padronizado para o menu."""
    return tk.Button(
        parent,
        text=text,
        command=command,
        bg=COR_AZUL_ESCURO,
        fg=COR_BRANCA,
        font=FONTE_TITULO,
        relief=tk.FLAT,
        height=2,
        width=25
    )


# CLASSE DE CARREGAMENTO E PROCESSAMENTO DE DADOS

class DataLoader:
    """
    Carrega, limpa e unifica os dados das planilhas.
    busca por Escola (Escola_Key) e cálculo de frequência.
    """

    def __init__(self):
        self.df_alunos = pd.DataFrame()
        self.df_presenca_completa = pd.DataFrame()
        self.is_loaded = False
        self.error_message = ""
        self.total_dias_por_oficina = {}

    def _load_trilhas_formativas(self):
        """Carrega e unifica os dados cadastrais, criando a chave normalizada de escola (Escola_Key)."""
        df_list = []
        if not os.path.exists(PLANILHA_TRILHAS):
            raise FileNotFoundError(f"Arquivo não encontrado: {PLANILHA_TRILHAS}")

        xls = pd.ExcelFile(PLANILHA_TRILHAS)
        col_map_lower = {k.lower(): v for k, v in COLUNAS_ALUNOS.items()}

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)

            df.columns = [col.lower().replace(' ', '_') for col in df.columns]
            df.rename(columns=col_map_lower, inplace=True)

            # Criação da chave para busca por escola
            school_name_key = normalize_text(sheet_name).replace(' ', '_').strip()
            school_name_display = sheet_name.replace('_', ' ').title()

            df['Escola'] = school_name_display
            df['Escola_Key'] = school_name_key  # Chave usada na busca

            if 'direcao' in df.columns:
                df['Direcao'] = df['direcao'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
                df['Direcao'] = df['Direcao'].fillna('')
            else:
                df['Direcao'] = ''

            df_list.append(df)

        df_alunos = pd.concat(df_list, ignore_index=True, sort=False)

        df_alunos['Matricula'] = df_alunos['Matricula'].astype(str).str.strip().str.upper()
        df_alunos.drop_duplicates(subset=['Matricula'], keep='first', inplace=True)

        df_alunos['Aluno_Normalized'] = df_alunos['Aluno'].apply(normalize_text)

        return df_alunos

    def _load_presenca_trilhas(self, df_alunos):
        """
        Carrega os dados de presença e unifica.
        """
        all_presences_list = []
        if not os.path.exists(PLANILHA_PRESENCA):
            raise FileNotFoundError(f"Arquivo não encontrado: {PLANILHA_PRESENCA}")

        xls = pd.ExcelFile(PLANILHA_PRESENCA)

        for sheet_name in xls.sheet_names:
            df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None)

            if df_raw.empty or len(df_raw.columns) < 2 or len(df_raw) < 2:
                continue

            office_name = normalize_text(sheet_name)
            office_name_title = sheet_name.replace('_', ' ').title()

            dates = df_raw.iloc[0, :].tolist()

            date_cols = [
                (i, col.strftime('%Y-%m-%d') if isinstance(col, pd.Timestamp) else str(col).strip())
                for i, col in enumerate(dates) if pd.notna(col) and str(col).strip() != ''
            ]

            dias_totais = len(date_cols)
            self.total_dias_por_oficina[office_name_title] = dias_totais

            for row_index in range(1, len(df_raw)):
                for col_index, date_str in date_cols:
                    present_student_name_raw = df_raw.iloc[row_index, col_index]

                    if pd.notna(present_student_name_raw) and str(present_student_name_raw).strip() != '':
                        normalized_present_name = normalize_text(present_student_name_raw)

                        all_presences_list.append({
                            'Aluno_Normalized': normalized_present_name,
                            'Oficina': office_name_title,
                            'Data_Oficina': date_str,
                            'Presenca': 1
                        })

        df_presenca_nomes = pd.DataFrame(all_presences_list)

        if df_presenca_nomes.empty:
            self.df_presenca_completa = pd.DataFrame()
            return self.df_presenca_completa

        # Remove duplicatas para garantir que cada aluno conte 1x por dia/oficina.
        df_presenca_nomes.drop_duplicates(subset=['Aluno_Normalized', 'Oficina', 'Data_Oficina'], inplace=True)

        alunos_com_matricula = df_alunos[
            ['Matricula', 'Escola', 'Escola_Key', 'Aluno', 'Aluno_Normalized']].copy().drop_duplicates(
            subset=['Aluno_Normalized', 'Matricula'])
        alunos_com_matricula = alunos_com_matricula[alunos_com_matricula['Aluno_Normalized'] != '']

        df_oficinas_info = pd.DataFrame(self.total_dias_por_oficina.items(), columns=['Oficina', 'Dias_Totais_Oficina'])

        from itertools import product
        all_combinations = list(
            product(alunos_com_matricula['Matricula'].unique(), df_oficinas_info['Oficina'].unique()))

        df_cartesian = pd.DataFrame(all_combinations, columns=['Matricula', 'Oficina'])
        df_cartesian = pd.merge(df_cartesian, df_oficinas_info, on='Oficina', how='left')

        df_base = pd.merge(
            df_cartesian,
            alunos_com_matricula.drop(columns=['Escola_Key']),
            on='Matricula',
            how='left'
        )

        self.df_presenca_completa = pd.merge(
            df_base,
            df_presenca_nomes[['Aluno_Normalized', 'Oficina', 'Data_Oficina', 'Presenca']],
            on=['Aluno_Normalized', 'Oficina'],
            how='left'
        )

        self.df_presenca_completa['Presenca'] = self.df_presenca_completa['Presenca'].fillna(0)
        self.df_presenca_completa['Data_Oficina'] = self.df_presenca_completa['Data_Oficina'].fillna('')

        return self.df_presenca_completa

    def load_data(self):
        """Ponto de entrada para carregar todos os dados."""
        if not os.path.exists(PLANILHA_TRILHAS) or not os.path.exists(PLANILHA_PRESENCA):
            self.error_message = ERRO_ARQUIVO_NAO_ENCONTRADO
            self.is_loaded = False
            return

        try:
            self.df_alunos = self._load_trilhas_formativas()
            alunos_validos = self.df_alunos[self.df_alunos['Aluno_Normalized'] != ''].copy()
            df_completo = self._load_presenca_trilhas(alunos_validos)

            if df_completo.empty and not self.df_alunos.empty:
                self.error_message = "Dados cadastrais carregados, mas a unificação de presença falhou (ou não há registros de presença)."
                self.is_loaded = True
                return
            elif df_completo.empty and self.df_alunos.empty:
                self.error_message = "Nenhum dado cadastral ou de presença carregado."
                self.is_loaded = False
                return

            self.df_presenca_completa = df_completo
            self.is_loaded = True

        except Exception as e:
            self.error_message = f"{ERRO_DADOS}\nDetalhe do erro: {e}"
            self.is_loaded = False
            print(f"Erro fatal no carregamento de dados: {e}")


# DEFINIÇÃO DAS CLASSES DE TELAS

class MenuFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_CINZA_CLARO)
        self.controller = controller

        try:
            # Carrega a imagem de fundo (usando o caminho configurado em Const.py)
            img = Image.open(IMAGEM_FUNDO)
            img = img.resize((JANELA_LARGURA, JANELA_ALTURA), Image.LANCZOS)
            self.bg_image = ImageTk.PhotoImage(img)

            bg_label = tk.Label(self, image=self.bg_image)
            bg_label.place(x=0, y=0, relwidth=1, relheight=1)

        except FileNotFoundError:
            tk.Label(self, text="Arquivo de imagem de fundo não encontrado.", font=FONTE_TITULO,
                     bg=COR_CINZA_CLARO).pack(pady=20)
        except Exception as e:
            tk.Label(self, text=f"Erro ao carregar imagem.", font=FONTE_TITULO, bg=COR_CINZA_CLARO).pack(pady=20)

        title = tk.Label(self, text="Análise de Trilhas Formativas", font=('Helvetica', 18, 'bold'), bg=COR_CINZA_CLARO)
        title.pack(pady=40)

        btn_busca = create_menu_button(self, "Busca de Dados Cadastrais",
                                       lambda: controller.show_frame("DadosAlunosFrame"))
        btn_busca.pack(pady=10)

        btn_frequencia = create_menu_button(self, "Cálculo de Frequência",
                                            lambda: controller.show_frame("PorcentagensFrame"))
        btn_frequencia.pack(pady=10)

        btn_sair = create_menu_button(self, "Sair", controller.quit)
        btn_sair.pack(pady=10)


class PorcentagensFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_CINZA_CLARO)
        self.controller = controller

        # Interface (Entradas)
        label = tk.Label(self, text="Cálculo de Frequência e Porcentagens", font=FONTE_TITULO, bg=COR_CINZA_CLARO)
        label.pack(pady=10)

        input_frame = tk.Frame(self, bg=COR_CINZA_CLARO)
        input_frame.pack(pady=10)

        tk.Label(input_frame, text="Buscar por:", bg=COR_CINZA_CLARO, font=FONTE_PRINCIPAL).pack(side=tk.LEFT, padx=5)

        self.search_by_var = tk.StringVar(self)
        self.search_by_var.set("Aluno")  # Valor padrão

        options = ["Aluno", "Matricula", "Oficina"]
        search_by_menu = ttk.Combobox(input_frame, textvariable=self.search_by_var, values=options, state="readonly",
                                      font=FONTE_PRINCIPAL)
        search_by_menu.pack(side=tk.LEFT, padx=5)

        self.search_entry = tk.Entry(input_frame, width=30, font=FONTE_PRINCIPAL)
        self.search_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(input_frame, text="Pesquisar", command=self.perform_search, bg=COR_AZUL_ESCURO, fg=COR_BRANCA,
                  relief=tk.FLAT).pack(side=tk.LEFT, padx=10)

        # Área de Resultados
        self.results_text = tk.Text(self, wrap=tk.WORD, width=70, height=25, font=FONTE_PRINCIPAL, bg=COR_BRANCA)
        self.results_text.pack(pady=10, padx=20)

        # Botão de retorno
        btn_voltar = create_menu_button(self, "Voltar ao Menu", lambda: controller.show_frame("MenuFrame"))
        btn_voltar.pack(pady=10)

    def calculate_percentage(self, df_filtered, search_value):
        """
        Calcula a frequência por oficina para um aluno (ou todos os alunos da oficina).
        """

        self.results_text.delete(1.0, tk.END)

        if df_filtered.empty:
            self.results_text.insert(tk.END, f"Nenhum registro encontrado para '{search_value}'.")
            return

        df = df_filtered.copy()

        # Agrupa por Matricula e Oficina para somar as presenças e contar os dias totais únicos
        df_group = df.groupby(['Matricula', 'Oficina', 'Dias_Totais_Oficina']).agg(
            Presencas_Contadas=('Presenca', 'sum'),
            Aluno=('Aluno', 'first'),
            Escola=('Escola', 'first')
        ).reset_index()

        df_group['Frequencia_Percentual'] = (
                    df_group['Presencas_Contadas'] / df_group['Dias_Totais_Oficina'] * 100).replace([np.inf, -np.inf],
                                                                                                    0).round(1)

        output = ""

        # Se a busca for por Aluno ou Matrícula
        if self.search_by_var.get() in ["Aluno", "Matricula"]:
            if df_group.empty:
                output += f"Nenhuma frequência de oficina registrada para '{search_value}'."
            else:
                aluno_info = df_group.iloc[0]
                output += f"--- Detalhes do Aluno: {aluno_info['Aluno']} ---\n"
                output += f"Escola: {aluno_info['Escola']}\n"
                output += f"Matrícula: {aluno_info['Matricula']}\n\n"
                output += "--- Frequência em Oficinas ---\n"

                for index, row in df_group.iterrows():
                    output += f"Oficina: {row['Oficina']} | Presença: {int(row['Presencas_Contadas'])}/{int(row['Dias_Totais_Oficina'])} ({row['Frequencia_Percentual']}%)"

                    dias_presentes = df[(df['Matricula'] == row['Matricula']) & (df['Oficina'] == row['Oficina']) & (
                                df['Presenca'] == 1)]['Data_Oficina'].unique()
                    if len(dias_presentes) > 0:
                        output += f"\nDias Presentes: {', '.join(dias_presentes)}\n"
                    else:
                        output += "\nDias Presentes: Nenhum dia registrado\n"

        # Se a busca for por Oficina
        elif self.search_by_var.get() == "Oficina":
            oficina_nome = df_group.iloc[0]['Oficina']
            dias_totais = df_group.iloc[0]['Dias_Totais_Oficina']

            output += f"--- Detalhes da Oficina: {oficina_nome} ---\n"
            output += f"Dias Totais da Oficina (Configurados): {dias_totais}\n\n"
            output += "--- Frequência Individual dos Alunos ---\n"

            for index, row in df_group.iterrows():
                output += f"Aluno: {row['Aluno']} (Escola: {row['Escola']}) | Frequência: {int(row['Presencas_Contadas'])}/{int(row['Dias_Totais_Oficina'])} ({row['Frequencia_Percentual']}%) \n"

            output += "\n--- Alunos Presentes por Dia (LISTA COMPLETA) ---\n"

            df_presenca_por_dia = df[df['Presenca'] == 1].groupby('Data_Oficina').agg(
                presentes=('Aluno', list),
                count=('Aluno', 'count')
            ).reset_index()

            # Exibe todos os nomes, sem limite.
            for index, row in df_presenca_por_dia.iterrows():
                output += f"Data {row['Data_Oficina']} ({row['count']} presentes)\n"
                output += ", ".join(row['presentes'])
                output += "\n"

        self.results_text.insert(tk.END, output)

    def perform_search(self):
        """Executa a busca e chama o cálculo de frequência."""
        search_value = self.search_entry.get().strip()
        search_by = self.search_by_var.get()

        if not search_value:
            messagebox.showwarning("Aviso", "Por favor, insira um valor para pesquisa.")
            return

        df_completo = self.controller.data_loader.df_presenca_completa

        if df_completo.empty:
            self.results_text.insert(tk.END, "Erro: Nenhum dado de presença carregado.")
            return

        normalized_search = normalize_text(search_value)
        df_filtered = pd.DataFrame()

        try:
            if search_by == "Aluno":
                df_filtered = df_completo[
                    df_completo['Aluno_Normalized'].str.contains(normalized_search, na=False)].copy()
            elif search_by == "Matricula":
                df_filtered = df_completo[
                    df_completo['Matricula'].astype(str).str.contains(normalized_search, na=False, case=False)].copy()
            elif search_by == "Oficina":
                df_filtered = df_completo[
                    df_completo['Oficina'].str.contains(search_value, na=False, case=False)].copy()

            self.calculate_percentage(df_filtered, search_value)

        except Exception as e:
            self.results_text.insert(tk.END, f"Ocorreu um erro inesperado durante a pesquisa: {e}")
            print(f"Erro na busca/cálculo de frequência: {e}")


class DadosAlunosFrame(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent, bg=COR_CINZA_CLARO)
        self.controller = controller

        # Interface (Entradas)
        label = tk.Label(self, text="Busca de Dados Cadastrais dos Alunos", font=FONTE_TITULO, bg=COR_CINZA_CLARO)
        label.pack(pady=10)

        input_frame = tk.Frame(self, bg=COR_CINZA_CLARO)
        input_frame.pack(pady=10)

        tk.Label(input_frame, text="Buscar por:", bg=COR_CINZA_CLARO, font=FONTE_PRINCIPAL).pack(side=tk.LEFT, padx=5)

        self.search_by_var = tk.StringVar(self)
        self.search_by_var.set("Aluno")  # Valor padrão

        options = CAMPOS_BUSCA_DADOS
        search_by_menu = ttk.Combobox(input_frame, textvariable=self.search_by_var, values=options, state="readonly",
                                      font=FONTE_PRINCIPAL)
        search_by_menu.pack(side=tk.LEFT, padx=5)

        self.search_entry = tk.Entry(input_frame, width=30, font=FONTE_PRINCIPAL)
        self.search_entry.pack(side=tk.LEFT, padx=5)

        tk.Button(input_frame, text="Pesquisar", command=self.perform_search, bg=COR_AZUL_ESCURO, fg=COR_BRANCA,
                  relief=tk.FLAT).pack(side=tk.LEFT, padx=10)

        # Área de Resultados
        self.results_text = tk.Text(self, wrap=tk.WORD, width=70, height=25, font=FONTE_PRINCIPAL, bg=COR_BRANCA)
        self.results_text.pack(pady=10, padx=20)

        # Botão de retorno
        btn_voltar = create_menu_button(self, "Voltar ao Menu", lambda: controller.show_frame("MenuFrame"))
        btn_voltar.pack(pady=10)

    def perform_search(self):
        """Executa a busca de dados cadastrais."""
        search_value = self.search_entry.get().strip()
        search_by = self.search_by_var.get()

        self.results_text.delete(1.0, tk.END)

        if not search_value:
            messagebox.showwarning("Aviso", "Por favor, insira um valor para pesquisa.")
            return

        df_alunos = self.controller.data_loader.df_alunos

        if df_alunos.empty:
            self.results_text.insert(tk.END, "Erro: Nenhum dado cadastral carregado.")
            return

        normalized_search = normalize_text(search_value)
        df_filtered = pd.DataFrame()

        try:
            if search_by == "Aluno":
                df_filtered = df_alunos[df_alunos['Aluno_Normalized'].str.contains(normalized_search, na=False)].copy()
            elif search_by == "Escola":
                df_filtered = df_alunos[
                    df_alunos['Escola_Key'].str.contains(normalized_search, na=False, case=False)].copy()
            else:
                df_filtered = df_alunos[
                    df_alunos[search_by].astype(str).apply(normalize_text).str.contains(normalized_search,
                                                                                        na=False)].copy()

            if df_filtered.empty:
                self.results_text.insert(tk.END,
                                         f"Nenhum aluno encontrado para '{search_value}' no campo '{search_by}'.")
                return

            output = f"--- Encontrados {len(df_filtered)} alunos com '{search_value}' ---\n\n"
            fields_to_show = ["Aluno", "Matricula", "CPF", "Mae", "Pai", "Turma", "Telefone", "Escola", "Direcao"]

            for index, row in df_filtered.iterrows():
                output += f"--- Aluno {index + 1} (Matrícula: {row['Matricula']})---\n"
                for field in fields_to_show:
                    if field in row:
                        output += f"{field}: {row[field]}\n"
                output += "\n"

            self.results_text.insert(tk.END, output)

        except Exception as e:
            self.results_text.insert(tk.END, f"Ocorreu um erro inesperado durante a pesquisa: {e}")
            print(f"Erro na busca de dados cadastrais: {e}")

# CLASSE PRINCIPAL DA APLICAÇÃO

class App(tk.Tk):
    def __init__(self, *args, **kwargs):
        # GARANTE A INICIALIZAÇÃO DA CLASSE TK
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("Análise de Trilhas Formativas")
        self.geometry(f"{JANELA_LARGURA}x{JANELA_ALTURA}")
        self.resizable(False, False)

        # Carregador de Dados
        self.data_loader = DataLoader()
        self.data_loader.load_data()

        # Container de Frames
        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        # Inicialização dos frames (janelas da aplicação)
        for F in (MenuFrame, PorcentagensFrame, DadosAlunosFrame):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        self.show_frame("MenuFrame")

        # Exibe o status do carregamento
        if not self.data_loader.is_loaded:
            messagebox.showwarning("Aviso",
                                   "A interface abriu, mas o carregamento dos dados falhou. Verifique os arquivos Excel.")

    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def quit(self):
        self.destroy()


# EXECUÇÃO DO APLICATIVO

if __name__ == "__main__":
    app = App()
    app.mainloop()
