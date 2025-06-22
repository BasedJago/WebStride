import tkinter as tk
from tkinter import messagebox, simpledialog, Listbox, Scrollbar, Frame, Label, Entry, Button, OptionMenu, StringVar, Radiobutton, Toplevel, filedialog, Text, Checkbutton, BooleanVar, ttk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException, NoSuchElementException, InvalidSessionIdException
from urllib3.exceptions import ProtocolError
import threading
import time
import os
import json
import logging
import re
import csv
from copy import deepcopy
from functools import partial
import webbrowser
from datetime import datetime
import pyautogui

# --- CLASSE PRINCIPAL DA APLICAÇÃO ---
class AutomationGUI:
    """
    Classe que constrói a interface gráfica e gerencia a lógica de automação.
    """
    def __init__(self, root):
        self.root = root
        self.root.title("WebStride Automatizador")
        self.root.geometry("800x600") 
        self.root.configure(bg="#2E2E2E")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.driver = None

        self.base_path = r"C:\WebStride"
        self.db_path = os.path.join(self.base_path, "Database")
        self.db_file = os.path.join(self.db_path, "database.json")
        self.log_file = os.path.join(self.base_path, "log.txt")
        self.error_log_file = os.path.join(self.base_path, "log_erro.txt")

        self.profiles_data = {}
        self.active_profile = ""
        self.profile_var = StringVar()
        
        self.selector_type_var = StringVar()
        self.SELECTOR_TYPES = ["CSS Selector", "XPath", "ID", "Nome da Classe", "Nome da Tag", "Texto Exato do Link", "Texto Parcial do Link", "Atributo name", "Texto Visível"]
        
        self.ACTIONS_REQUIRING_SELECTOR = [
            "Clicar em Elemento", "Clicar e Esperar", "Escrever em Campo", 
            "Selecionar Tudo (Ctrl+A)", "Copiar (Ctrl+C)", "Colar (Ctrl+V)", "Upload de Arquivo",
            "Extrair Texto de Elemento", "Extrair Atributo de Elemento", "Mover Mouse para Elemento (Hover)",
            "Rolar até Elemento", "Duplo Clique em Elemento", 
            "Extrair Tabela para Arquivo CSV"
        ]

        self.imported_data = None
        self.imported_data_path_var = StringVar(value="Nenhum arquivo de dados carregado.")
        self.raw_data_content = ""
        
        self.headless_var = BooleanVar(value=False)
        self.drag_start_index = None
        self.internal_variables = {}
        
        self.dialog_result = None
        self.dialog_event = threading.Event()
        self.test_driver = None

        # --- NOVA ESTRUTURA DE CATEGORIAS DE AÇÕES ---
        self.action_categories = {
            "Navegador e Página": [
                "Abrir Site",
                "Atualizar Página (F5)",
                "Navegar (Voltar)",
                "Navegar (Avançar)",
                "Aguardar Fechamento do Navegador",
                "Trocar para Aba",
                "Fechar Aba",
                "Tirar Print da Tela",
                "Executar JavaScript",
            ],
            "Ações Globais (PyAutoGUI)": [
                "Clicar em Coordenadas (X,Y)",
                "Clicar em Coordenadas e Escrever",
                "Pressionar Múltiplas Teclas",
                "Pressionar Enter",
                "Digitar Texto (Global)",
            ],
            "Interação com Elementos (Selenium)": [
                "Clicar em Elemento",
                "Clicar e Esperar",
                "Duplo Clique em Elemento",
                "Mover Mouse para Elemento (Hover)",
                "Escrever em Campo",
                "Upload de Arquivo",
                "Rolar até Elemento",
                "Selecionar Tudo (Ctrl+A)",
                "Copiar (Ctrl+C)",
                "Colar (Ctrl+V)",
            ],
            "Extração de Dados (Selenium)": [
                "Extrair Texto de Elemento",
                "Extrair Atributo de Elemento",
                "Extrair Tabela para Arquivo CSV",
            ],
            "Controle de Fluxo e Lógica": [
                "Aguardar (segundos)",
                "Esperar Verificação Humana",
                "Iniciar Loop",
                "Iniciar Loop Fixo",
                "Fim do Loop",
                "Fim do Loop e Perguntar",
                "Se (condição)",
                "Senão Se (condição)",
                "Senão",
                "Fim Se",
                "Separador Visual",
            ],
            "Variáveis e Área de Transferência": [
                "Criar Variável Interna",
                "Pedir Input do Usuário",
                "Manipular Variável",
                "Salvar Área de Transferência em Variável",
                "Salvar Variável na Área de Transferência",
            ],
        }

        self.ACTION_MAP = {
            # Core
            "Abrir Site": self.act_open_site,
            "Aguardar (segundos)": self.act_wait,
            # Lógica
            "Iniciar Loop": self.act_pass,
            "Iniciar Loop Fixo": self.act_pass,
            "Fim do Loop": self.act_pass, 
            "Fim do Loop e Perguntar": self.act_pass,
            "Se (condição)": self.act_pass,
            "Senão Se (condição)": self.act_pass,
            "Senão": self.act_pass,
            "Fim Se": self.act_pass,
            "Esperar Verificação Humana": self.act_wait_for_human,
            "Separador Visual": self.act_pass,
            # Ações de Elemento
            "Clicar em Elemento": self.act_click_element,
            "Clicar e Esperar": self.act_click_and_wait,
            "Escrever em Campo": self.act_type_in_element,
            "Upload de Arquivo": self.act_upload_file,
            # Ações de Teclado/Clipboard
            "Pressionar Enter": self.act_press_enter,
            "Copiar (Ctrl+C)": self.act_copy,
            "Colar (Ctrl+V)": self.act_paste,
            "Selecionar Tudo (Ctrl+A)": self.act_select_all,
            "Pressionar Múltiplas Teclas": self.act_press_multiple_keys,
            "Digitar Texto (Global)": self.act_type_globally,
            "Salvar Área de Transferência em Variável": self.act_save_clipboard_to_variable,
            "Salvar Variável na Área de Transferência": self.act_save_variable_to_clipboard, # --- NOVO ---
            # Ações de Navegador
            "Trocar para Aba": self.act_switch_tab,
            "Fechar Aba": self.act_close_tab,
            # Ações de Variáveis e Input
            "Criar Variável Interna": self.act_create_internal_variable,
            "Pedir Input do Usuário": self.act_prompt_user_input,
            # Ações de Extração
            "Extrair Texto de Elemento": self.act_extract_text,
            "Extrair Atributo de Elemento": self.act_extract_attribute,
            # Ações de Interação
            "Mover Mouse para Elemento (Hover)": self.act_hover_element,
            "Rolar a Página": self.act_scroll_page,
            "Rolar até Elemento": self.act_scroll_to_element,
            "Duplo Clique em Elemento": self.act_double_click,
            "Clicar em Coordenadas (X,Y)": self.act_click_at_xy,
            "Clicar em Coordenadas e Escrever": self.act_click_at_xy_and_type,
            # Ações de Controle/Navegador
            "Tirar Print da Tela": self.act_take_screenshot,
            "Executar JavaScript": self.act_execute_js,
            "Atualizar Página (F5)": self.act_refresh_page,
            "Navegar (Voltar)": self.act_navigate_back,
            "Navegar (Avançar)": self.act_navigate_forward,
            "Aguardar Fechamento do Navegador": self.act_wait_for_browser_close,
            # Ações de Dados
            "Manipular Variável": self.act_manipulate_variable,
            "Extrair Tabela para Arquivo CSV": self.act_extract_table_to_csv,
        }
        
        self.flat_action_list = sorted(list(self.ACTION_MAP.keys()))
        
        self.categorized_action_list = []
        for category, actions in self.action_categories.items():
            self.categorized_action_list.append(f"--- {category.upper()} ---")
            self.categorized_action_list.extend(sorted(actions))

        self._setup_logging()
        self._setup_styles()
        self._create_widgets()      
        self._setup_data_storage()  
        self.load_initial_profile() 

    def _execute_dialog_on_main_thread(self, dialog_func_with_args):
        self.dialog_result = dialog_func_with_args()
        self.dialog_event.set()

    def _setup_logging(self):
        if not os.path.exists(self.base_path): os.makedirs(self.base_path)
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        logger = logging.getLogger()
        if not logger.handlers:
            info_handler = logging.FileHandler(self.log_file, encoding='utf-8')
            info_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
            error_handler = logging.FileHandler(self.error_log_file, encoding='utf-8')
            error_handler.setLevel(logging.ERROR)
            error_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
            logger.addHandler(info_handler)
            logger.addHandler(error_handler)
        logging.info("Aplicação iniciada.")

    def _setup_data_storage(self):
        try:
            if not os.path.exists(self.db_path): os.makedirs(self.db_path)
            self.load_data_from_json()
        except OSError as e:
            messagebox.showerror("Erro de Diretório", f"Não foi possível criar o diretório {self.db_path}\nErro: {e}")
            logging.error(f"Falha ao criar diretório: {e}")
            self.root.quit()

    def _setup_styles(self):
        self.style = {
            "bg": "#2E2E2E", "fg": "#FFFFFF", "entry_bg": "#3C3C3C",
            "button_bg": "#5E5E5E", "button_active_bg": "#7A7A7A",
            "listbox_bg": "#3C3C3C", "accent_color": "#007ACC", "link_color": "#6CB4EE",
            "separator_color": "#9A9A9A" 
        }

    def _create_widgets(self):
        header_frame = Frame(self.root, bg="#1E1E1E")
        header_frame.pack(fill=tk.X, side=tk.TOP)
        Label(header_frame, text="WebStride Automatizador", font=("Arial", 12, "bold"), bg="#1E1E1E", fg=self.style["fg"]).pack(side=tk.LEFT, padx=10, pady=5)
        Button(header_frame, text="?", font=("Arial", 10, "bold"), command=self.show_tutorial, bg=self.style["accent_color"], fg=self.style["fg"], relief=tk.FLAT, width=3).pack(side=tk.RIGHT, padx=10, pady=5)
        Button(header_frame, text="⚙", font=("Arial", 12), command=self.open_settings_window, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT, width=3).pack(side=tk.RIGHT, padx=5, pady=5)

        profile_frame = Frame(self.root, bg="#1E1E1E", padx=10, pady=5)
        profile_frame.pack(fill=tk.X, side=tk.TOP)
        Label(profile_frame, text="Gerenciar Perfis:", bg="#1E1E1E", fg=self.style["fg"]).pack(side=tk.LEFT, padx=5)
        self.profile_menu = OptionMenu(profile_frame, self.profile_var, "")
        self.profile_menu.config(bg=self.style["button_bg"], fg=self.style["fg"], activebackground=self.style["button_active_bg"], relief=tk.FLAT)
        self.profile_menu.pack(side=tk.LEFT, padx=5)
        Button(profile_frame, text="Novo", command=self.create_new_profile, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(profile_frame, text="Renomear", command=self.rename_profile, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(profile_frame, text="Excluir", command=self.delete_profile, bg="#C82333", fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        
        data_import_frame = Frame(self.root, bg="#252526", padx=10, pady=5)
        data_import_frame.pack(fill=tk.X, side=tk.TOP)
        Button(data_import_frame, text="Importar Dados", command=self.import_data_file, bg=self.style["accent_color"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        data_path_label = Label(data_import_frame, textvariable=self.imported_data_path_var, bg="#252526", fg=self.style["link_color"], cursor="hand2")
        data_path_label.pack(side=tk.LEFT, padx=10)
        data_path_label.bind("<Button-1>", self.view_data_table)

        top_frame = Frame(self.root, bg=self.style["bg"], padx=10, pady=10)
        top_frame.pack(fill=tk.X, side=tk.TOP)

        browser_frame = Frame(top_frame, bg=self.style["bg"])
        browser_frame.pack(side=tk.LEFT, padx=(0, 20), anchor='n')
        Label(browser_frame, text="Navegador:", bg=self.style["bg"], fg=self.style["fg"]).pack(anchor='w')
        self.browser_var = StringVar(value="Chrome")
        Radiobutton(browser_frame, text="Chrome", variable=self.browser_var, value="Chrome", bg=self.style["bg"], fg=self.style["fg"], selectcolor=self.style["entry_bg"], activebackground=self.style["bg"]).pack(anchor='w')
        Radiobutton(browser_frame, text="Firefox", variable=self.browser_var, value="Firefox", bg=self.style["bg"], fg=self.style["fg"], selectcolor=self.style["entry_bg"], activebackground=self.style["bg"]).pack(anchor='w')
        Checkbutton(browser_frame, text="Executar em modo invisível (headless)", variable=self.headless_var, bg=self.style["bg"], fg=self.style["fg"], selectcolor=self.style["entry_bg"], activebackground=self.style["bg"], activeforeground=self.style["fg"]).pack(anchor='w')

        action_builder_frame = Frame(top_frame, bg=self.style["bg"])
        action_builder_frame.pack(side=tk.LEFT, expand=True, fill=tk.X)
        action_builder_frame.grid_columnconfigure(1, weight=2)

        Label(action_builder_frame, text="Ação:", bg=self.style["bg"], fg=self.style["fg"]).grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.action_var = StringVar(value=list(self.ACTION_MAP.keys())[0])
        
        # Permite que o usuário digite no campo (state='normal')
        self.action_menu = ttk.Combobox(action_builder_frame, textvariable=self.action_var, state='normal', font=("Arial", 9))
        # Define os valores iniciais usando a lista que criamos no Passo 1
        self.action_menu['values'] = self.categorized_action_list
        self.action_menu.grid(row=0, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
        # Conecta a digitação a uma nova função de filtro
        self.action_menu.bind('<KeyRelease>', self._update_action_combobox_filter)
        
        self.action_var.trace_add("write", self._on_action_change)

        Label(action_builder_frame, text="Nome/Descrição (Opcional):", bg=self.style["bg"], fg=self.style["fg"]).grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.desc_entry = Entry(action_builder_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        self.desc_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
        
        self.selector_type_label = Label(action_builder_frame, text="Tipo de Seletor:", bg=self.style["bg"], fg=self.style["fg"])
        self.selector_type_var.set(self.SELECTOR_TYPES[0])
        self.selector_type_menu = OptionMenu(action_builder_frame, self.selector_type_var, *self.SELECTOR_TYPES)
        self.selector_type_menu.config(bg=self.style["button_bg"], fg=self.style["fg"], activebackground=self.style["button_active_bg"], relief=tk.FLAT)
        
        self.param_label = Label(action_builder_frame, text="Parâmetro:", bg=self.style["bg"], fg=self.style["fg"])
        self.param_entry = Entry(action_builder_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        
        self.test_selector_button = Button(action_builder_frame, text="Testar", command=self.test_selector, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT)
        
        self.param2_label = Label(action_builder_frame, text="Parâmetro 2:", bg=self.style["bg"], fg=self.style["fg"])
        self.param2_entry = Entry(action_builder_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)

        self.param3_label = Label(action_builder_frame, text="Parâmetro 3:", bg=self.style["bg"], fg=self.style["fg"])
        self.param3_entry = Entry(action_builder_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        
        self.param_label.grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.param_entry.grid(row=3, column=1, padx=5, pady=2, sticky='ew')
        
        self._on_action_change()

        Button(action_builder_frame, text="Adicionar Ação", command=self.add_action, bg=self.style["accent_color"], fg=self.style["fg"], relief=tk.FLAT).grid(row=7, column=1, columnspan=2, padx=5, pady=10, sticky='e')

        bottom_frame = Frame(self.root, bg=self.style["bg"], padx=10, pady=10)
        bottom_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        tools_frame = Frame(self.root, bg="#252526", padx=10, pady=5)
        tools_frame.pack(fill=tk.X, side=tk.TOP, before=bottom_frame)

        Label(tools_frame, text="Ferramentas de Desenvolvedor:", bg="#252526", fg=self.style["fg"]).pack(side=tk.LEFT, padx=5)
        
        Button(tools_frame, text="Abrir Navegador de Teste", command=self.open_test_browser, bg="#007ACC", fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(tools_frame, text="Fechar Navegador de Teste", command=self.close_test_browser, bg="#5E5E5E", fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        list_controls_frame = Frame(bottom_frame, bg=self.style["bg"])
        list_controls_frame.pack(side=tk.LEFT)
        Button(list_controls_frame, text="Editar Ação", command=self.edit_selected_action, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(list_controls_frame, text="Indentar ->", command=self.indent_selection, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(list_controls_frame, text="<- Recuar", command=self.unindent_selection, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(list_controls_frame, text="Remover Ação", command=self.remove_selected_action, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)
        Button(list_controls_frame, text="Limpar Ações", command=self.clear_all_actions, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.LEFT, padx=5)

        exec_controls_frame = Frame(bottom_frame, bg=self.style["bg"])
        exec_controls_frame.pack(side=tk.RIGHT)
        Button(exec_controls_frame, text="Salvar Perfis", command=self.save_data_to_json, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)
        self.exec_button = Button(exec_controls_frame, text="▶ Executar", command=self.start_automation_thread, bg="#28A745", fg=self.style["fg"], relief=tk.FLAT)
        self.exec_button.pack(side=tk.RIGHT, padx=5)
        Button(exec_controls_frame, text="■ Parar Navegador", command=self.stop_automation, bg="#DC3545", fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)

        variables_frame = Frame(self.root, bg=self.style["bg"], padx=10, pady=5)
        variables_frame.pack(fill=tk.X, side=tk.BOTTOM)
        Label(variables_frame, text="Variáveis Internas (em tempo real):", bg=self.style["bg"], fg=self.style["fg"]).pack(anchor='w')

        tree_frame = Frame(variables_frame)
        tree_frame.pack(fill=tk.X, expand=True)
        style = ttk.Style(self.root)
        style.configure("Treeview", background=self.style["listbox_bg"], foreground=self.style["fg"], fieldbackground=self.style["listbox_bg"], borderwidth=0)
        style.map('Treeview', background=[('selected', self.style["accent_color"])])
        
        style.configure("TCombobox",
                        fieldbackground=self.style["entry_bg"],
                        background=self.style["button_bg"],
                        foreground=self.style["fg"],
                        arrowcolor=self.style["fg"],
                        bordercolor=self.style["separator_color"],
                        lightcolor=self.style["entry_bg"],
                        darkcolor=self.style["entry_bg"])

        style.map('TCombobox',
                  # Força o fundo do campo e a cor da fonte em todos os estados não desabilitados
                  fieldbackground=[('!disabled', self.style["entry_bg"])],
                  # --- LINHA MODIFICADA AQUI ---
                  foreground=[('!disabled', 'black')],
                  # Cor de fundo do item selecionado na lista dropdown (azul)
                  selectbackground=[('!disabled', self.style["accent_color"])],
                  # Cor da fonte do item selecionado na lista dropdown (branca)
                  selectforeground=[('!disabled', self.style["fg"])])
        
        self.variable_treeview = ttk.Treeview(tree_frame, columns=("Variable", "Value"), show="headings", height=4)
        self.variable_treeview.heading("Variable", text="Variável")
        self.variable_treeview.heading("Value", text="Valor")
        self.variable_treeview.column("Variable", width=150)
        self.variable_treeview.pack(side=tk.LEFT, fill=tk.X, expand=True)
        var_scrollbar = Scrollbar(tree_frame, orient="vertical", command=self.variable_treeview.yview)
        var_scrollbar.pack(side=tk.RIGHT, fill="y")
        self.variable_treeview.config(yscrollcommand=var_scrollbar.set)
        
        list_frame = Frame(self.root, bg=self.style["bg"], padx=10, pady=10)
        list_frame.pack(fill=tk.BOTH, expand=True, side=tk.TOP)
        self.listbox_label = Label(list_frame, text="Sequência de Ações (Macro):", bg=self.style["bg"], fg=self.style["fg"])
        self.listbox_label.pack(anchor='w')
        self.listbox = Listbox(list_frame, bg=self.style["listbox_bg"], fg=self.style["fg"], selectbackground=self.style["accent_color"], relief=tk.FLAT, height=15)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = Scrollbar(list_frame, orient="vertical", command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)
        self.listbox.bind("<Button-1>", self.on_drag_start)
        self.listbox.bind("<B1-Motion>", self.on_drag_motion)
        self.listbox.bind("<ButtonRelease-1>", self.on_drop)
        self.listbox.bind("<Double-1>", self.edit_selected_action)

    def open_test_browser(self):
        if self.test_driver:
            messagebox.showinfo("Informação", "O navegador de teste já está aberto.")
            return
        try:
            logging.info("Abrindo navegador de teste.")
            chrome_options = webdriver.ChromeOptions()
            self.test_driver = webdriver.Chrome(options=chrome_options)
            self.test_driver.maximize_window()
            self.test_driver.get("https://www.google.com")
        except Exception as e:
            messagebox.showerror("Erro ao Abrir Navegador", f"Não foi possível iniciar o navegador de teste.\nVerifique se o ChromeDriver está atualizado.\n\nErro: {e}")
            logging.error(f"Falha ao abrir navegador de teste: {e}")
            self.test_driver = None

    def close_test_browser(self):
        if self.test_driver:
            logging.info("Fechando navegador de teste.")
            try:
                self.test_driver.quit()
            except Exception as e:
                logging.error(f"Erro ao fechar o navegador de teste: {e}")
            finally:
                self.test_driver = None
        else:
            messagebox.showinfo("Informação", "O navegador de teste já está fechado.")
            
    def _ensure_driver(self):
        """Verifica se o driver do navegador existe e está ativo. Levanta uma exceção se não estiver."""
        if not self.driver or not self.driver.window_handles:
            raise Exception("Ação de navegador executada sem um navegador aberto ou com o navegador fechado. Use 'Abrir Site' no início da sua macro.")

    def _find_element(self, action_data, timeout=10):
        self._ensure_driver() 
        selector_type = action_data.get('selector_type', 'CSS Selector')
        selector_value = action_data.get('param', '')
        
        if selector_type == "Texto Visível":
            by_method = By.XPATH
            if "'" in selector_value and '"' in selector_value:
                parts = selector_value.split("'")
                xpath_text = "concat('" + "', \"'\", '".join(parts) + "')"
                selector_value = f"//*[contains(text(), {xpath_text})]"
            elif "'" in selector_value:
                selector_value = f'//*[contains(text(), "{selector_value}")]'
            else:
                selector_value = f"//*[contains(text(), '{selector_value}')]"
        else:
            by_map = {
                "CSS Selector": By.CSS_SELECTOR, "XPath": By.XPATH, "ID": By.ID, 
                "Nome da Classe": By.CLASS_NAME, "Nome da Tag": By.TAG_NAME,
                "Texto Exato do Link": By.LINK_TEXT, "Texto Parcial do Link": By.PARTIAL_LINK_TEXT,
                "Atributo name": By.NAME
            }
            by_method = by_map.get(selector_type, By.CSS_SELECTOR)
            
        return WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((by_method, selector_value)))
        
    def _find_element_for_test(self, driver, action_data, timeout=10):
        """Versão do _find_element que aceita um driver como argumento."""
        selector_type = action_data.get('selector_type', 'CSS Selector')
        selector_value = action_data.get('param', '')
        
        if selector_type == "Texto Visível":
            by_method = By.XPATH
            if "'" in selector_value and '"' in selector_value:
                parts = selector_value.split("'")
                xpath_text = "concat('" + "', \"'\", '".join(parts) + "')"
                selector_value = f"//*[contains(text(), {xpath_text})]"
            elif "'" in selector_value:
                selector_value = f'//*[contains(text(), "{selector_value}")]'
            else:
                selector_value = f"//*[contains(text(), '{selector_value}')]"
        else:
            by_map = {
                "CSS Selector": By.CSS_SELECTOR, "XPath": By.XPATH, "ID": By.ID, 
                "Nome da Classe": By.CLASS_NAME, "Nome da Tag": By.TAG_NAME,
                "Texto Exato do Link": By.LINK_TEXT, "Texto Parcial do Link": By.PARTIAL_LINK_TEXT,
                "Atributo name": By.NAME
            }
            by_method = by_map.get(selector_type, By.CSS_SELECTOR)
            
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by_method, selector_value)))

    def on_closing(self):
        logging.info("Aplicação fechada. Salvando dados.")
        self.save_data_to_json()
        self.close_test_browser()
        self.stop_automation()
        self.root.destroy()
        
    def show_tutorial(self):
        try:
            webbrowser.open('http://webstridetutorial.blogspot.com')
        except Exception as e:
            logging.error(f"Erro ao abrir o tutorial: {e}")
            messagebox.showerror("Erro", f"Não foi possível abrir o link do tutorial.\n\n{e}")

    def open_settings_window(self):
        # TO-DO: Carregar as configurações salvas (do database.json, por exemplo)
        
        settings_window = Toplevel(self.root)
        settings_window.title("Configurações - WebStride")
        settings_window.geometry("600x400")
        settings_window.configure(bg=self.style["bg"])
        settings_window.transient(self.root) # Mantém a janela de config na frente da principal
        settings_window.grab_set() # Bloqueia interação com a janela principal

        main_frame = Frame(settings_window, bg=self.style["bg"], padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_columnconfigure(1, weight=1)

        # --- CAMINHO BASE ---
        Label(main_frame, text="Diretório Base:", bg=self.style["bg"], fg=self.style["fg"]).grid(row=0, column=0, padx=5, pady=8, sticky='w')
        base_path_entry = Entry(main_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        base_path_entry.grid(row=0, column=1, padx=5, pady=8, sticky='ew')
        # TO-DO: Inserir valor atual da configuração aqui
        # base_path_entry.insert(0, self.config.get('base_path', 'C:\\WebStride'))

        # --- TIMEOUT PADRÃO ---
        Label(main_frame, text="Timeout Padrão (segundos):", bg=self.style["bg"], fg=self.style["fg"]).grid(row=1, column=0, padx=5, pady=8, sticky='w')
        timeout_entry = Entry(main_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        timeout_entry.grid(row=1, column=1, padx=5, pady=8, sticky='ew')
        # TO-DO: Inserir valor atual da configuração aqui
        # timeout_entry.insert(0, self.config.get('default_timeout', '10'))

        # --- CAMINHO DO WEBDRIVER ---
        Label(main_frame, text="Caminho do ChromeDriver:", bg=self.style["bg"], fg=self.style["fg"]).grid(row=2, column=0, padx=5, pady=8, sticky='w')
        driver_path_entry = Entry(main_frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"], relief=tk.FLAT)
        driver_path_entry.grid(row=2, column=1, padx=5, pady=8, sticky='ew')
        # TO-DO: Inserir valor atual da configuração aqui
        # driver_path_entry.insert(0, self.config.get('webdriver_path', ''))

        # --- BOTÕES DE AÇÃO ---
        button_frame = Frame(main_frame, bg=self.style["bg"])
        button_frame.grid(row=3, column=0, columnspan=2, pady=20, sticky='e')

        def save_settings():
            # TO-DO: Lógica para salvar os valores dos campos Entry
            # Ex: self.config['base_path'] = base_path_entry.get()
            # self.save_data_to_json() # Salvar o json que agora conteria os perfis e as configs
            messagebox.showinfo("Salvo", "Configurações salvas com sucesso!", parent=settings_window)
            settings_window.destroy()

        Button(button_frame, text="Salvar e Fechar", command=save_settings, bg=self.style["accent_color"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.RIGHT, padx=5)
        Button(button_frame, text="Cancelar", command=settings_window.destroy, bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT).pack(side=tk.RIGHT)
       
    def import_data_file(self):
        filepath = filedialog.askopenfilename(title="Selecionar arquivo de dados", filetypes=(("Arquivos de Dados", "*.txt *.csv"), ("Todos os arquivos", "*.*")))
        if not filepath: return
        if self._load_data_from_path(filepath):
            if self.active_profile:
                self.profiles_data["profiles"][self.active_profile]["data_file_path"] = filepath
                logging.info(f"Caminho do arquivo '{filepath}' salvo para o perfil '{self.active_profile}'.")

    def _load_data_from_path(self, filepath):
        try:
            self.raw_data_content = ""
            rows = []
            with open(filepath, 'r', encoding='utf-8') as f:
                self.raw_data_content = f.read()
                f.seek(0)
                if filepath.lower().endswith('.csv'):
                    reader = csv.reader(f, delimiter=';')
                    headers = next(reader)
                    for row in reader:
                        rows.append(row)
                else: 
                    lines = f.readlines()
                    if not lines: raise ValueError("O arquivo de dados está vazio.")
                    headers = [h.strip() for h in lines[0].split(';')]
                    rows = [[cell.strip() for cell in line.split(';')] for line in lines[1:]]
            
            self.imported_data = {"headers": headers, "rows": rows}
            self.imported_data_path_var.set(filepath)
            logging.info(f"Arquivo de dados carregado com sucesso: {filepath}")
            return True
        except Exception as e:
            self._clear_imported_data()
            logging.error(f"Falha ao importar dados de {filepath}: {e}")
            messagebox.showerror("Erro de Importação", f"Não foi possível ler o arquivo de dados.\nVerifique o formato e o caminho.\nErro: {e}")
            return False
            
    def _clear_imported_data(self):
        self.imported_data = None
        self.raw_data_content = ""
        self.imported_data_path_var.set("Nenhum arquivo de dados carregado.")

    def view_data_table(self, event=None):
        if not self.imported_data:
            messagebox.showinfo("Nenhum Dado", "Nenhum arquivo de dados foi carregado ainda.")
            return
            
        viewer = Toplevel(self.root)
        viewer.title(f"Visualizador de Dados - {os.path.basename(self.imported_data_path_var.get())}")
        viewer.geometry("800x500")
        viewer.configure(bg=self.style["bg"])
        
        style = ttk.Style(viewer)
        style.theme_use("clam")
        style.configure("Treeview", background=self.style["listbox_bg"], foreground=self.style["fg"], fieldbackground=self.style["listbox_bg"], borderwidth=0)
        style.map('Treeview', background=[('selected', self.style["accent_color"])])
        
        tree_frame = Frame(viewer)
        tree_frame.pack(pady=5, padx=10, fill=tk.BOTH, expand=True)

        tree = ttk.Treeview(tree_frame, columns=self.imported_data["headers"], show='headings')
        
        for header in self.imported_data["headers"]:
            tree.heading(header, text=header)
            tree.column(header, anchor="w", width=120)

        for i, row in enumerate(self.imported_data["rows"]):
            tree.insert('', tk.END, iid=i, values=row)
            
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        
        status_bar = Label(viewer, text="Clique em uma célula para copiar a variável correspondente.", bd=1, relief=tk.SUNKEN, anchor=tk.W, bg=self.style["bg"], fg=self.style["fg"])
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        def on_cell_click(event):
            region = tree.identify("region", event.x, event.y)
            if region != "cell":
                return

            row_id = tree.identify_row(event.y)
            column_id = tree.identify_column(event.x)
            
            if not row_id or not column_id:
                return

            try:
                col_index = int(column_id.replace('#', '')) - 1
                col_name = self.imported_data["headers"][col_index]
                row_num = int(row_id) + 1
                
                variable_syntax = f"${col_name}{row_num}"
                
                self.root.clipboard_clear()
                self.root.clipboard_append(variable_syntax)
                
                status_bar.config(text=f"Variável '{variable_syntax}' copiada para a área de transferência!")
                viewer.after(3000, lambda: status_bar.config(text="Clique em uma célula para copiar a variável correspondente."))
            except (IndexError, ValueError) as e:
                logging.warning(f"Não foi possível obter a variável da célula clicada: {e}")
                status_bar.config(text="Não foi possível obter a variável desta célula.")

        tree.bind("<Button-1>", on_cell_click)

    def _resolve_variable(self, text_param, loop_iteration=None):
        if not isinstance(text_param, str):
            return text_param

        def replace_internal_match(match):
            var_name = match.group(1)
            return str(self.internal_variables.get(var_name, match.group(0)))

        internal_pattern = r'#([a-zA-Z_][a-zA-Z0-9_]*)'
        resolved_text = re.sub(internal_pattern, replace_internal_match, text_param)

        if self.imported_data:
            def replace_data_match(match):
                col_name = match.group(1)
                row_num_str = match.group(2)
                if not row_num_str: return match.group(0)
                row_num = int(row_num_str)

                effective_row_num = row_num
                if loop_iteration is not None:
                    effective_row_num = row_num + loop_iteration
                
                try:
                    col_index = self.imported_data["headers"].index(col_name)
                    row_index = effective_row_num - 1 
                    if 0 <= row_index < len(self.imported_data["rows"]) and 0 <= col_index < len(self.imported_data["rows"][row_index]):
                        return self.imported_data["rows"][row_index][col_index]
                    else:
                        # --- CORRIGIDO --- Linha do SyntaxError
                        logging.warning(f"Índice de dados inválido para ${col_name}{effective_row_num}.")
                        return match.group(0)
                except (ValueError, IndexError):
                    logging.warning(f"Coluna '{col_name}' ou linha base '{row_num}' não encontrada.")
                    return match.group(0)

            data_pattern = r'\$([a-zA-Z_][a-zA-Z0-9_]*)(\d+)'
            resolved_text = re.sub(data_pattern, replace_data_match, resolved_text)

        return resolved_text

    def load_initial_profile(self):
        last_profile = self.profiles_data.get("last_open_profile")
        profiles = list(self.profiles_data.get("profiles", {}).keys())
        if last_profile and last_profile in profiles: self.profile_var.set(last_profile)
        elif profiles: self.profile_var.set(sorted(profiles)[0])
        else: self.create_new_profile("Perfil Padrão", supress_dialog=True)
        self.switch_profile()

    def load_data_from_json(self):
        try:
            if os.path.exists(self.db_file):
                with open(self.db_file, 'r', encoding='utf-8') as f: self.profiles_data = json.load(f)
                for name, data in self.profiles_data.get("profiles", {}).items():
                    if isinstance(data, list): self.profiles_data["profiles"][name] = {"actions": data, "data_file_path": ""}
                    profile_actions = self.profiles_data["profiles"][name].get("actions", [])
                    for action in profile_actions:
                        if "indent" not in action: action["indent"] = 0
                        if "description" not in action: action["description"] = ""
                        if "param3" not in action: action["param3"] = ""
            else: self.profiles_data = {"profiles": {}, "last_open_profile": ""}
            self._update_profile_menu()
        except (json.JSONDecodeError, IOError) as e:
            logging.error(f"Erro ao carregar JSON: {e}")
            self.profiles_data = {"profiles": {}, "last_open_profile": ""}
            self._update_profile_menu()

    def save_data_to_json(self):
        try:
            if self.active_profile: self.profiles_data["last_open_profile"] = self.active_profile
            with open(self.db_file, 'w', encoding='utf-8') as f: json.dump(self.profiles_data, f, indent=4)
            logging.info(f"Perfis salvos em {self.db_file}")
        except Exception as e:
            logging.error(f"Erro ao salvar JSON: {e}")

    def _update_profile_menu(self):
        self.action_menu['values'] = sorted(list(self.ACTION_MAP.keys()))
        
        profile_menu = self.profile_menu["menu"]; profile_menu.delete(0, "end")
        profiles = self.profiles_data.get("profiles", {}).keys()
        if not profiles: self.profile_var.set("")
        else:
            for profile in sorted(profiles): profile_menu.add_command(label=profile, command=lambda v=profile: (self.profile_var.set(v), self.switch_profile()))
            if not self.profile_var.get() in profiles: self.profile_var.set(sorted(profiles)[0] if profiles else "")

    def switch_profile(self, *args):
        new_profile = self.profile_var.get()
        if new_profile:
            self.active_profile = new_profile
            profile_data = self.profiles_data["profiles"].get(self.active_profile, {})
            data_path = profile_data.get("data_file_path", "")
            if data_path and os.path.exists(data_path): self._load_data_from_path(data_path)
            else: self._clear_imported_data()
            self.update_listbox()
            logging.info(f"Perfil trocado para: {self.active_profile}")
        self.listbox_label.config(text=f"Sequência de Ações do Perfil: '{self.active_profile}'")

    def create_new_profile(self, name=None, supress_dialog=False):
        new_name = name
        if not supress_dialog: new_name = simpledialog.askstring("Novo Perfil", "Digite o nome do novo perfil:")
        if new_name and new_name not in self.profiles_data["profiles"]:
            self.profiles_data["profiles"][new_name] = {"actions": [], "data_file_path": ""}
            self._update_profile_menu(); self.profile_var.set(new_name); self.switch_profile()
            logging.info(f"Novo perfil criado: {new_name}")
        elif new_name and not supress_dialog: messagebox.showwarning("Nome Existente", "Já existe um perfil com este nome.")

    def rename_profile(self):
        if not self.active_profile: return
        new_name = simpledialog.askstring("Renomear Perfil", f"Digite o novo nome para '{self.active_profile}':", initialvalue=self.active_profile)
        if new_name and new_name != self.active_profile and new_name not in self.profiles_data["profiles"]:
            old_name = self.active_profile
            self.profiles_data["profiles"][new_name] = self.profiles_data["profiles"].pop(old_name)
            self.active_profile = new_name
            self._update_profile_menu(); self.profile_var.set(new_name); self.switch_profile()
            logging.info(f"Perfil '{old_name}' renomeado para '{new_name}'")

    def delete_profile(self):
        if not self.active_profile: return
        if messagebox.askyesno("Excluir Perfil", f"Tem certeza que deseja excluir o perfil '{self.active_profile}'?"):
            old_name = self.active_profile
            del self.profiles_data["profiles"][old_name]
            logging.info(f"Perfil excluído: {old_name}")
            self.active_profile = ""
            self._update_profile_menu(); self.load_initial_profile()
            
    def _on_action_change(self, *args):
        # --- MODIFICADO ---
        action = self.action_var.get()
        self.selector_type_label.grid_remove(); self.selector_type_menu.grid_remove()
        self.param_label.grid_remove(); self.param_entry.grid_remove()
        self.param2_label.grid_remove(); self.param2_entry.grid_remove()
        self.param3_label.grid_remove(); self.param3_entry.grid_remove()
        self.test_selector_button.grid_remove()
        
        no_param_actions = ["Pressionar Enter", "Esperar Verificação Humana", "Iniciar Loop", "Fim do Loop", 
                            "Senão", "Fim Se", "Fechar Aba", "Fim do Loop e Perguntar", "Atualizar Página (F5)",
                            "Navegar (Voltar)", "Navegar (Avançar)", "Aguardar Fechamento do Navegador"]
        actions_with_condition = ["Se (condição)", "Senão Se (condição)"]
        
        if action in self.ACTIONS_REQUIRING_SELECTOR:
            self.selector_type_label.grid(row=2, column=0, padx=5, pady=2, sticky='w')
            self.selector_type_menu.grid(row=2, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
            self.param_label.config(text="Seletor:")
            self.param_label.grid(row=3, column=0, padx=5, pady=2, sticky='w')
            self.param_entry.grid(row=3, column=1, padx=5, pady=2, sticky='ew')
            self.test_selector_button.grid(row=3, column=2, padx=5, pady=2, sticky='w')

        elif action not in no_param_actions:
            param_text = "Parâmetro:"
            if action in actions_with_condition: param_text = "Condição:"
            elif action == "Salvar Área de Transferência em Variável": param_text = "Variável de Destino:"
            elif action == "Salvar Variável na Área de Transferência": param_text = "Variável de Origem (ex: #var):"
            elif action == "Separador Visual": param_text = "Título do Separador:"
            elif action == "Rolar a Página": param_text = "Direção/Pixels:"
            elif action == "Tirar Print da Tela": param_text = "Nome do Arquivo (Opc):"
            elif action == "Executar JavaScript": param_text = "Script JS:"
            elif action == "Criar Variável Interna": param_text = "Nome da Variável:"
            elif action == "Pedir Input do Usuário": param_text = "Mensagem para Usuário:"
            elif action == "Manipular Variável": param_text = "Variável de Destino:"
            elif action == "Pressionar Múltiplas Teclas": param_text = "Teclas (ex: ctrl+t):"
            elif action == "Clicar em Coordenadas (X,Y)": param_text = "Coordenada X:"
            elif action == "Clicar em Coordenadas e Escrever": param_text = "Coordenada X:"
            elif action == "Extrair Tabela para Arquivo CSV": self.param2_label.config(text="Nome do Arquivo CSV:")
            elif action == "Digitar Texto (Global)": param_text = "Texto a Digitar:"

            self.param_label.config(text=param_text)
            self.param_label.grid(row=3, column=0, padx=5, pady=2, sticky='w')
            self.param_entry.grid(row=3, column=1, columnspan=2, padx=5, pady=2, sticky='ew')

        if action in ["Escrever em Campo", "Upload de Arquivo", "Criar Variável Interna", "Pedir Input do Usuário", "Extrair Texto de Elemento", "Manipular Variável", "Clicar em Coordenadas (X,Y)"]:
            self.param2_label.grid(row=4, column=0, padx=5, pady=2, sticky='w')
            self.param2_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
            if action == "Upload de Arquivo": self.param2_label.config(text="Caminho do Arquivo:")
            elif action == "Escrever em Campo": self.param2_label.config(text="Valor:")
            elif action == "Criar Variável Interna": self.param2_label.config(text="Valor Inicial:")
            elif action == "Pedir Input do Usuário": self.param2_label.config(text="Salvar na Variável:")
            elif action == "Extrair Texto de Elemento": self.param2_label.config(text="Salvar na Variável:")
            elif action == "Manipular Variável": self.param2_label.config(text="Expressão (ex: #v1+10):")
            elif action == "Clicar em Coordenadas (X,Y)": self.param2_label.config(text="Coordenada Y:")

        if action == "Extrair Atributo de Elemento":
            self.param2_label.config(text="Nome do Atributo:")
            self.param2_label.grid(row=4, column=0, padx=5, pady=2, sticky='w')
            self.param2_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
            self.param3_label.config(text="Salvar na Variável:")
            self.param3_label.grid(row=5, column=0, padx=5, pady=2, sticky='w')
            self.param3_entry.grid(row=5, column=1, columnspan=2, padx=5, pady=2, sticky='ew')

        if action == "Clicar em Coordenadas e Escrever":
            self.param2_label.config(text="Coordenada Y:")
            self.param2_label.grid(row=4, column=0, padx=5, pady=2, sticky='w')
            self.param2_entry.grid(row=4, column=1, columnspan=2, padx=5, pady=2, sticky='ew')
            self.param3_label.config(text="Texto a Escrever:")
            self.param3_label.grid(row=5, column=0, padx=5, pady=2, sticky='w')
            self.param3_entry.grid(row=5, column=1, columnspan=2, padx=5, pady=2, sticky='ew')

    # CÓDIGO A SER SUBSTITUÍDO
    def _update_action_combobox_filter(self, event=None):
        """Filtra a lista de ações e abre o dropdown, ignorando os separadores."""
        cursor_pos = self.action_menu.index(tk.INSERT)
        typed_text = self.action_var.get()
        
        # Ignora teclas que não devem acionar o filtro
        if event and event.keysym in ('Down', 'Up', 'Return', 'Escape', 'Tab', 'Right', 'Left', 'Home', 'End'):
            # Se a lista estiver filtrada e o usuário pressionar Esc, restaura a lista completa
            if event.keysym == 'Escape':
                self.action_menu['values'] = self.categorized_action_list
            return

        # Se não houver texto digitado, exibe a lista completa com categorias
        if not typed_text:
            self.action_menu['values'] = self.categorized_action_list
            return

        # Se houver texto, filtra a lista de ações "reais" (sem os separadores)
        filtered_list = [
            action for action in self.flat_action_list 
            if typed_text.lower() in action.lower()
        ]
        
        # Atualiza o menu com os resultados filtrados
        self.action_menu['values'] = filtered_list
        self.action_menu.icursor(cursor_pos)

        # Força a abertura do menu se houver resultados
        if filtered_list and (not event or event.keysym != 'BackSpace'):
            # Atraso mínimo para garantir que a UI atualizou os valores antes de abrir
            self.action_menu.after(50, lambda: self.action_menu.event_generate('<Down>'))

    def test_selector(self):
        active_driver = self.test_driver if self.test_driver and self.test_driver.window_handles else self.driver
        
        if not active_driver or not active_driver.window_handles:
            messagebox.showinfo("Navegador Fechado", "Abra o 'Navegador de Teste' ou execute uma ação primeiro para testar um seletor.")
            return
            
        selector_type = self.selector_type_var.get()
        param = self.param_entry.get()
        if not param: messagebox.showwarning("Parâmetro Vazio", "Digite um seletor para testar."); return
        
        action_data = {"selector_type": selector_type, "param": param}
        
        def highlight(driver, element):
            original_style = driver.execute_script("return arguments[0].getAttribute('style');", element)
            driver.execute_script("arguments[0].setAttribute('style', 'border: 3px solid red; background-color: yellow;');", element)
            time.sleep(1.5)
            driver.execute_script("arguments[0].setAttribute('style', arguments[1]);", element, original_style)

        try:
            element = self._find_element_for_test(active_driver, action_data, timeout=5)
            highlight(active_driver, element)
            messagebox.showinfo("Sucesso", f"Elemento encontrado com sucesso usando [{selector_type}]: {param}")
        except Exception as e:
            messagebox.showerror("Elemento não Encontrado", f"Não foi possível encontrar o elemento.\nErro: {e}")

    def add_action(self):
        if not self.active_profile: return
        action_name = self.action_var.get()
        param = self.param_entry.get()
        param2 = self.param2_entry.get()
        param3 = self.param3_entry.get()
        description = self.desc_entry.get()
        
        new_action = {"action": action_name, "param": param, "param2": param2, "param3": param3, "indent": 0, "description": description}

        if action_name in self.ACTIONS_REQUIRING_SELECTOR:
            new_action["selector_type"] = self.selector_type_var.get()
            if not param: messagebox.showerror("Parâmetro Vazio", "Ações que usam seletor não podem ter o parâmetro vazio."); return
        
        self.profiles_data["profiles"][self.active_profile]["actions"].append(new_action)
        self.update_listbox()
        self.param_entry.delete(0, tk.END); self.param2_entry.delete(0, tk.END)
        self.param3_entry.delete(0, tk.END); self.desc_entry.delete(0, tk.END)
        logging.info(f"Ação '{action_name}' adicionada ao perfil '{self.active_profile}'.")

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        if self.active_profile:
            actions = self.profiles_data["profiles"].get(self.active_profile, {}).get("actions", [])
            numbering_stack = []
            
            actions_with_two_params = [
                "Escrever em Campo", "Upload de Arquivo", "Criar Variável Interna", "Pedir Input do Usuário",
                "Extrair Texto de Elemento", "Manipular Variável", "Extrair Tabela para Arquivo CSV"
            ]
            actions_with_three_params = ["Extrair Atributo de Elemento"]

            for i, action_data in enumerate(actions):
                action = action_data['action']
                indent_level = action_data.get('indent', 0)
                
                if action == "Separador Visual":
                    title = action_data.get('param', '').upper()
                    separator_text = f"{'  ' * indent_level}────────── {title} ──────────"
                    self.listbox.insert(tk.END, separator_text)
                    self.listbox.itemconfig(tk.END, {'fg': self.style["separator_color"]})
                    continue 

                while len(numbering_stack) > indent_level: numbering_stack.pop()
                if len(numbering_stack) <= indent_level:
                    numbering_stack.append(1)
                else:
                    numbering_stack[-1] +=1

                display_number = ".".join(map(str, numbering_stack))

                description = action_data.get('description', '')
                param = action_data.get('param', ''); sel_type = action_data.get('selector_type')
                indent_space = "  " * indent_level
                
                display_text = f"{indent_space}{display_number}. "
                if description:
                    display_text += f"[{description}] "
                
                display_text += action
                
                if param or sel_type: display_text += ": "
                
                if action == "Clicar em Coordenadas (X,Y)":
                    param2_val = action_data.get('param2', '')
                    display_text += f"X: {param} | Y: {param2_val}"
                elif action == "Clicar em Coordenadas e Escrever":
                    param2_val = action_data.get('param2', '')
                    param3_val = action_data.get('param3', '')
                    display_text += f"X: {param} | Y: {param2_val} | Texto: {param3_val}"
                else:
                    if sel_type: display_text += f"[{sel_type}] {param}"
                    elif param: display_text += param
                
                if action in actions_with_two_params:
                    param2_val = action_data.get('param2', '')
                    label = "Param2"
                    if action == "Upload de Arquivo": label = "Caminho"
                    elif action == "Escrever em Campo": label = "Valor"
                    elif action == "Criar Variável Interna": label = "Valor Inicial"
                    elif action == "Pedir Input do Usuário": label = "Salvar em"
                    elif action == "Extrair Texto de Elemento": label = "Salvar em"
                    elif action == "Manipular Variável": label = "Expressão"
                    elif action == "Extrair Tabela para Arquivo CSV": label = "Arquivo CSV"
                    display_text += f" | {label}: {param2_val}"

                if action in actions_with_three_params:
                    param2_val = action_data.get('param2', '')
                    param3_val = action_data.get('param3', '')
                    if action == "Extrair Atributo de Elemento":
                        display_text += f" | Atributo: {param2_val} | Salvar em: {param3_val}"

                self.listbox.insert(tk.END, display_text)

    def indent_selection(self):
        if not self.active_profile: return
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        actions = self.profiles_data["profiles"][self.active_profile]["actions"]
        for i in selected_indices:
             actions[i]["indent"] = min(actions[i].get("indent", 0) + 1, 5)
        self.update_listbox()

    def unindent_selection(self):
        if not self.active_profile: return
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        actions = self.profiles_data["profiles"][self.active_profile]["actions"]
        for i in selected_indices:
            if actions[i].get("indent", 0) > 0: actions[i]["indent"] -= 1
        self.update_listbox()
        
    def remove_selected_action(self):
        if not self.active_profile: return
        selected_indices = self.listbox.curselection()
        if not selected_indices: return
        for i in sorted(selected_indices, reverse=True): del self.profiles_data["profiles"][self.active_profile]["actions"][i]
        self.update_listbox()
        logging.info(f"Ação removida do perfil '{self.active_profile}'.")
        
    def edit_selected_action(self, event=None):
        if not self.active_profile: return
        selected_indices = self.listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("Nenhuma Seleção", "Selecione uma ação para editar.")
            return

        index = selected_indices[0]
        action_to_edit = deepcopy(self.profiles_data["profiles"][self.active_profile]["actions"][index])

        editor_window = Toplevel(self.root)
        editor_window.title("Editar Ação")
        editor_window.geometry("500x350")
        editor_window.configure(bg=self.style["bg"])

        edit_action_var = StringVar(value=action_to_edit.get("action"))
        edit_selector_type_var = StringVar(value=action_to_edit.get("selector_type", self.SELECTOR_TYPES[0]))
        
        frame = Frame(editor_window, bg=self.style["bg"], padx=10, pady=10)
        frame.pack(expand=True, fill=tk.BOTH)
        frame.grid_columnconfigure(1, weight=1)

        Label(frame, text="Ação:", bg=self.style["bg"], fg=self.style["fg"]).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        edit_action_menu = OptionMenu(frame, edit_action_var, *sorted(self.ACTION_MAP.keys()))
        edit_action_menu.config(bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT)
        edit_action_menu.grid(row=0, column=1, padx=5, pady=5, sticky='ew')

        Label(frame, text="Nome/Descrição:", bg=self.style["bg"], fg=self.style["fg"]).grid(row=1, column=0, padx=5, pady=5, sticky='w')
        edit_desc_entry = Entry(frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"])
        edit_desc_entry.insert(0, action_to_edit.get("description", ""))
        edit_desc_entry.grid(row=1, column=1, padx=5, pady=5, sticky='ew')

        edit_selector_type_label = Label(frame, text="Tipo de Seletor:", bg=self.style["bg"], fg=self.style["fg"])
        edit_selector_type_menu = OptionMenu(frame, edit_selector_type_var, *self.SELECTOR_TYPES)
        edit_selector_type_menu.config(bg=self.style["button_bg"], fg=self.style["fg"], relief=tk.FLAT)

        edit_param_label = Label(frame, text="Parâmetro:", bg=self.style["bg"], fg=self.style["fg"])
        edit_param_entry = Entry(frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"])
        edit_param_entry.insert(0, action_to_edit.get("param", ""))

        edit_param2_label = Label(frame, text="Parâmetro 2:", bg=self.style["bg"], fg=self.style["fg"])
        edit_param2_entry = Entry(frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"])
        edit_param2_entry.insert(0, action_to_edit.get("param2", ""))
        
        edit_param3_label = Label(frame, text="Parâmetro 3:", bg=self.style["bg"], fg=self.style["fg"])
        edit_param3_entry = Entry(frame, bg=self.style["entry_bg"], fg=self.style["fg"], insertbackground=self.style["fg"])
        edit_param3_entry.insert(0, action_to_edit.get("param3", ""))

        def _update_editor_ui(*args):
            # --- MODIFICADO ---
            action = edit_action_var.get()
            edit_selector_type_label.grid_remove(); edit_selector_type_menu.grid_remove()
            edit_param_label.grid_remove(); edit_param_entry.grid_remove()
            edit_param2_label.grid_remove(); edit_param2_entry.grid_remove()
            edit_param3_label.grid_remove(); edit_param3_entry.grid_remove()
            
            no_param_actions = ["Pressionar Enter", "Esperar Verificação Humana", "Iniciar Loop", "Fim do Loop", 
                                "Senão", "Fim Se", "Fechar Aba", "Fim do Loop e Perguntar", "Atualizar Página (F5)",
                                "Navegar (Voltar)", "Navegar (Avançar)", "Aguardar Fechamento do Navegador"]
            actions_with_condition = ["Se (condição)", "Senão Se (condição)"]
            
            if action in self.ACTIONS_REQUIRING_SELECTOR:
                edit_selector_type_label.grid(row=2, column=0, padx=5, pady=5, sticky='w')
                edit_selector_type_menu.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
                edit_param_label.config(text="Seletor:")
                edit_param_label.grid(row=3, column=0, padx=5, pady=5, sticky='w')
                edit_param_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
            elif action not in no_param_actions:
                param_text = "Parâmetro:"
                if action in actions_with_condition: param_text = "Condição:"
                elif action == "Salvar Área de Transferência em Variável": param_text = "Variável de Destino:"
                elif action == "Salvar Variável na Área de Transferência": param_text = "Variável de Origem (ex: #var):"
                elif action == "Separador Visual": param_text = "Título do Separador:"
                elif action == "Rolar a Página": param_text = "Direção/Pixels:"
                elif action == "Tirar Print da Tela": param_text = "Nome do Arquivo (Opc):"
                elif action == "Executar JavaScript": param_text = "Script JS:"
                elif action == "Criar Variável Interna": param_text = "Nome da Variável:"
                elif action == "Pedir Input do Usuário": param_text = "Mensagem para Usuário:"
                elif action == "Manipular Variável": param_text = "Variável de Destino:"
                elif action == "Pressionar Múltiplas Teclas": param_text = "Teclas (ex: ctrl+t):"
                elif action == "Clicar em Coordenadas (X,Y)": param_text = "Coordenada X:"
                elif action == "Clicar em Coordenadas e Escrever": param_text = "Coordenada X:"
                elif action == "Digitar Texto (Global)": param_text = "Texto a Digitar:"
                edit_param_label.config(text=param_text)
                edit_param_label.grid(row=3, column=0, padx=5, pady=5, sticky='w')
                edit_param_entry.grid(row=3, column=1, padx=5, pady=5, sticky='ew')
            
            if action in ["Escrever em Campo", "Upload de Arquivo", "Criar Variável Interna", "Pedir Input do Usuário", "Extrair Texto de Elemento", "Manipular Variável", "Clicar em Coordenadas (X,Y)"]:
                edit_param2_label.grid(row=4, column=0, padx=5, pady=5, sticky='w')
                edit_param2_entry.grid(row=4, column=1, padx=5, pady=5, sticky='ew')
                if action == "Upload de Arquivo": edit_param2_label.config(text="Caminho do Arquivo:")
                elif action == "Escrever em Campo": edit_param2_label.config(text="Valor:")
                elif action == "Criar Variável Interna": edit_param2_label.config(text="Valor Inicial:")
                elif action == "Pedir Input do Usuário": edit_param2_label.config(text="Salvar na Variável:")
                elif action == "Extrair Texto de Elemento": edit_param2_label.config(text="Salvar na Variável:")
                elif action == "Manipular Variável": edit_param2_label.config(text="Expressão (ex: #v1+10):")
                elif action == "Clicar em Coordenadas (X,Y)": edit_param2_label.config(text="Coordenada Y:")
                elif action == "Extrair Tabela para Arquivo CSV": edit_param2_label.config(text="Nome do Arquivo CSV:")
            
            if action == "Extrair Atributo de Elemento":
                edit_param2_label.config(text="Nome do Atributo:")
                edit_param2_label.grid(row=4, column=0, padx=5, pady=5, sticky='w')
                edit_param2_entry.grid(row=4, column=1, padx=5, pady=5, sticky='ew')
                edit_param3_label.config(text="Salvar na Variável:")
                edit_param3_label.grid(row=5, column=0, padx=5, pady=5, sticky='w')
                edit_param3_entry.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

            if action == "Clicar em Coordenadas e Escrever":
                edit_param2_label.config(text="Coordenada Y:")
                edit_param2_label.grid(row=4, column=0, padx=5, pady=5, sticky='w')
                edit_param2_entry.grid(row=4, column=1, padx=5, pady=5, sticky='ew')
                edit_param3_label.config(text="Texto a Escrever:")
                edit_param3_label.grid(row=5, column=0, padx=5, pady=5, sticky='w')
                edit_param3_entry.grid(row=5, column=1, padx=5, pady=5, sticky='ew')

        edit_action_var.trace_add("write", _update_editor_ui)
        _update_editor_ui() 

        def save_changes():
            new_action_data = {
                "action": edit_action_var.get(),
                "description": edit_desc_entry.get(),
                "param": edit_param_entry.get(),
                "selector_type": edit_selector_type_var.get(),
                "param2": edit_param2_entry.get(),
                "param3": edit_param3_entry.get(),
                "indent": action_to_edit.get("indent", 0) 
            }
            self.profiles_data["profiles"][self.active_profile]["actions"][index] = new_action_data
            self.update_listbox()
            logging.info(f"Ação {index + 1} do perfil '{self.active_profile}' foi editada.")
            editor_window.destroy()

        Button(frame, text="Salvar Alterações", command=save_changes, bg=self.style["accent_color"], fg=self.style["fg"], relief=tk.FLAT).grid(row=6, column=1, pady=10, sticky='e')

    def clear_all_actions(self):
        if not self.active_profile: return
        if messagebox.askyesno("Confirmar", f"Limpar TODAS as ações do perfil '{self.active_profile}'?"):
            self.profiles_data["profiles"][self.active_profile]["actions"] = []
            self.update_listbox()
            logging.warning(f"Todas as ações do perfil '{self.active_profile}' foram limpas.")

    def start_automation_thread(self):
        if not self.active_profile: return
        actions = self.profiles_data["profiles"].get(self.active_profile, {}).get("actions", [])
        if not actions: return
        self.exec_button.config(state=tk.DISABLED, text="Executando...")
        threading.Thread(target=self.run_automation_logic, daemon=True).start()
    
    def run_automation_logic(self):
        self.internal_variables = {}
        self.driver = None 
        self.root.after(0, self._update_variable_display)
        logging.info(f"Iniciando automação do perfil: {self.active_profile}")
        
        actions_to_run = list(self.profiles_data["profiles"].get(self.active_profile, {}).get("actions", []))
        try:
            exec_stack = [] 
            pc = 0
            
            while pc < len(actions_to_run):
                self.listbox.selection_clear(0, tk.END); self.listbox.selection_set(pc)
                action_data = actions_to_run[pc]
                action_name = action_data.get("action")

                if exec_stack and exec_stack[-1]['type'] in ['if', 'elif'] and not exec_stack[-1]['condition_met']:
                    if action_name in ["Senão Se (condição)", "Senão", "Fim Se"]: pass 
                    else: pc += 1; continue

                if action_name == "Iniciar Loop" or action_name == "Iniciar Loop Fixo":
                    loop_start_pc, end_loop_pc = self._find_block_end(actions_to_run, pc)
                    if end_loop_pc == -1: pc += 1; continue
                    if action_name == "Iniciar Loop Fixo": num_loops = int(self._resolve_variable(action_data.get("param", "0")))
                    else: num_loops = len(self.imported_data["rows"]) if self.imported_data else 0
                    exec_stack.append({'type': 'loop', 'start_pc': loop_start_pc, 'end_pc': end_loop_pc, 'total_iter': num_loops, 'current_iter': 0})
                    logging.info(f"Iniciando loop com {num_loops} iterações.")
                    pc += 1

                elif action_name == "Fim do Loop":
                    if exec_stack and exec_stack[-1]['type'] == 'loop':
                        loop_data = exec_stack[-1]
                        loop_data['current_iter'] += 1
                        if loop_data['current_iter'] < loop_data['total_iter']:
                            pc = loop_data['start_pc'] + 1
                            logging.info(f"--- Loop Iteração {loop_data['current_iter'] + 1}/{loop_data['total_iter']} ---")
                        else:
                            logging.info("Fim do loop.")
                            exec_stack.pop()
                            pc += 1
                    else:
                        pc += 1
                
                elif action_name == "Fim do Loop e Perguntar":
                    if exec_stack and exec_stack[-1]['type'] == 'loop':
                        loop_data = exec_stack[-1]
                        
                        self.dialog_event.clear()
                        dialog_call = partial(messagebox.askyesno, "Fim da Iteração", "Deseja executar o loop novamente desde o início?")
                        self.root.after(0, self._execute_dialog_on_main_thread, dialog_call)
                        self.dialog_event.wait()
                        response = self.dialog_result

                        if response:
                            loop_data['current_iter'] = 0
                            pc = loop_data['start_pc'] + 1
                            logging.info("Usuário escolheu reiniciar o loop.")
                        else:
                            exec_stack.pop()
                            pc += 1
                            logging.info("Usuário escolheu encerrar o loop.")
                    else:
                        pc += 1
                
                elif action_name == "Se (condição)":
                    cond_start_pc, cond_end_pc = self._find_block_end(actions_to_run, pc)
                    condition_str = self._resolve_variable(action_data.get("param", ""), self._get_current_loop_iter(exec_stack))
                    result = self._evaluate_condition(condition_str)
                    exec_stack.append({'type': 'if', 'start_pc': cond_start_pc, 'end_pc': cond_end_pc, 'condition_met': result, 'has_executed_if': result})
                    logging.info(f"Condição 'Se': '{condition_str}' avaliada como {result}")
                    if not result: pc = self._find_next_elif_else_or_endif(actions_to_run, pc)
                    else: pc += 1

                elif action_name == "Senão Se (condição)":
                    if exec_stack and exec_stack[-1]['type'] in ['if', 'elif']:
                        if exec_stack[-1]['has_executed_if']: pc = exec_stack[-1]['end_pc']
                        else:
                            condition_str = self._resolve_variable(action_data.get("param", ""), self._get_current_loop_iter(exec_stack))
                            result = self._evaluate_condition(condition_str)
                            exec_stack[-1]['type'] = 'elif'; exec_stack[-1]['condition_met'] = result; exec_stack[-1]['has_executed_if'] = result
                            logging.info(f"Condição 'Senão Se': '{condition_str}' avaliada como {result}")
                            if not result: pc = self._find_next_elif_else_or_endif(actions_to_run, pc)
                            else: pc += 1
                    else: pc += 1

                elif action_name == "Senão":
                    if exec_stack and exec_stack[-1]['type'] in ['if', 'elif']:
                        if exec_stack[-1]['has_executed_if']: pc = exec_stack[-1]['end_pc']
                        else:
                            exec_stack[-1]['condition_met'] = True; exec_stack[-1]['has_executed_if'] = True
                            logging.info("Entrando no bloco 'Senão'."); pc += 1
                    else: pc += 1
                
                elif action_name == "Fim Se":
                    if exec_stack and exec_stack[-1]['type'] in ['if', 'elif']: exec_stack.pop()
                    pc += 1
                
                else: 
                    loop_iter = self._get_current_loop_iter(exec_stack)
                    resolved_action = self._resolve_action_variables(action_data, loop_iter)
                    logging.info(f"Executando ação: {resolved_action.get('action')}")
                    if resolved_action.get("action") in self.ACTION_MAP:
                        self.ACTION_MAP[resolved_action.get("action")](resolved_action)
                    pc += 1

            self._write_data_to_txt_if_needed()
            self.root.after(0, messagebox.showinfo, "Sucesso", f"Automação do perfil '{self.active_profile}' concluída!")
            logging.info(f"Automação do perfil '{self.active_profile}' concluída.")

        except (TimeoutException, NoSuchElementException) as e:
            error_message = f"Elemento não encontrado ou demorou demais para aparecer (Timeout). Verifique o seletor na ação {pc + 1}.\n\nErro Original: {e}"
            logging.error(error_message, exc_info=True)
            self.root.after(0, messagebox.showerror, "Erro de Seletor", error_message)
        except (InvalidSessionIdException, ProtocolError, ConnectionRefusedError) as e:
            error_message = f"A conexão com o navegador foi perdida ou o navegador foi fechado inesperadamente.\n\nErro Original: {e}"
            logging.error(error_message, exc_info=True)
            self.root.after(0, messagebox.showerror, "Erro de Conexão com Navegador", error_message)
        except UserCancelledException:
            self.root.after(0, messagebox.showinfo, "Cancelado", "A automação foi cancelada pelo usuário.")
            logging.info("Automação cancelada pelo usuário durante verificação humana.")
        except Exception as e:
            error_message = f"Ocorreu um erro inesperado na automação (Perfil: {self.active_profile}):\n\nAção: {action_data.get('action') if 'action_data' in locals() else 'N/A'} \nErro: {e}"
            logging.error(error_message, exc_info=True)
            self.root.after(0, messagebox.showerror, "Erro na Automação", error_message)
        finally:
            self.root.after(0, self.reset_exec_button)
            self.stop_automation()
    
    def _get_current_loop_iter(self, exec_stack):
        for item in reversed(exec_stack):
            if item['type'] == 'loop': return item['current_iter']
        return None

    def _find_block_end(self, actions, start_pc):
        start_indent = actions[start_pc].get("indent", 0)
        block_type = actions[start_pc].get("action")
        end_action_map = {
            "Iniciar Loop": ["Fim do Loop", "Fim do Loop e Perguntar"], 
            "Iniciar Loop Fixo": ["Fim do Loop", "Fim do Loop e Perguntar"], 
            "Se (condição)": ["Fim Se"]
        }
        end_action_names = end_action_map.get(block_type)
        if not end_action_names: return start_pc, -1

        for i in range(start_pc + 1, len(actions)):
            action = actions[i]
            if action.get("indent", 0) == start_indent and action.get("action") in end_action_names:
                return start_pc, i
        return start_pc, -1

    def _find_next_elif_else_or_endif(self, actions, current_pc):
        start_indent = actions[current_pc].get("indent", 0)
        for i in range(current_pc + 1, len(actions)):
            action = actions[i]
            if action.get("indent", 0) == start_indent:
                if action.get("action") in ["Senão Se (condição)", "Senão", "Fim Se"]: return i
            elif action.get("indent", 0) < start_indent: break
        return len(actions) 

    def _evaluate_condition(self, condition_str):
        try:
            parts = []
            if '==' in condition_str: parts = [p.strip() for p in condition_str.split('==', 1)] + ['==']
            elif '!=' in condition_str: parts = [p.strip() for p in condition_str.split('!=', 1)] + ['!=']
            elif '>=' in condition_str: parts = [p.strip() for p in condition_str.split('>=', 1)] + ['>=']
            elif '<=' in condition_str: parts = [p.strip() for p in condition_str.split('<=', 1)] + ['<=']
            elif '>' in condition_str: parts = [p.strip() for p in condition_str.split('>', 1)] + ['>']
            elif '<' in condition_str: parts = [p.strip() for p in condition_str.split('<', 1)] + ['<']
            elif ' não contém ' in condition_str: parts = [p.strip() for p in condition_str.split(' não contém ', 1)] + ['não contém']
            elif ' contém ' in condition_str: parts = [p.strip() for p in condition_str.split(' contém ', 1)] + ['contém']
            else: return False
            var, val, op = parts
            try: num_val = float(val); num_var = float(var)
            except ValueError: num_val = val; num_var = var
            if op == '==': return num_var == num_val
            if op == '!=': return num_var != num_val
            if op == '>': return isinstance(num_var, float) and num_var > num_val
            if op == '<': return isinstance(num_var, float) and num_var < num_val
            if op == '>=': return isinstance(num_var, float) and num_var >= num_val
            if op == '<=': return isinstance(num_var, float) and num_var <= num_val
            if op == 'contém': return str(val) in str(var)
            if op == 'não contém': return str(val) not in str(var)
        except Exception as e:
            logging.error(f"Erro ao avaliar condição '{condition_str}': {e}")
            return False
        return False

    def _resolve_action_variables(self, action_data, loop_iteration):
        resolved_data = deepcopy(action_data)
        for key, value in resolved_data.items():
            if isinstance(value, str):
                resolved_data[key] = self._resolve_variable(value, loop_iteration)
        return resolved_data

    def _write_data_to_txt_if_needed(self):
        current_path = self.imported_data_path_var.get()
        if not self.imported_data or not os.path.exists(current_path): return
        try:
            output_lines = [";".join(self.imported_data["headers"])]
            for row in self.imported_data["rows"]:
                output_lines.append(";".join(map(str, row)))
            content = "\n".join(output_lines)
            if content.strip() != self.raw_data_content.strip():
                with open(current_path, 'w', encoding='utf-8', newline='') as f:
                    if current_path.lower().endswith('.csv'):
                        writer = csv.writer(f, delimiter=';'); writer.writerow(self.imported_data["headers"]); writer.writerows(self.imported_data["rows"])
                    else: f.write(content)
                logging.info(f"Arquivo de dados '{current_path}' foi atualizado com sucesso.")
                self.raw_data_content = content
        except Exception as e:
            logging.error(f"Falha ao salvar o arquivo de dados de volta para '{current_path}': {e}")
            messagebox.showwarning("Aviso", "Não foi possível salvar as alterações no arquivo de dados.")

    def reset_exec_button(self): self.exec_button.config(state=tk.NORMAL, text="▶ Executar")
    def stop_automation(self):
        if self.driver: self.driver.quit(); self.driver = None; logging.info("Navegador fechado.")

    def _update_variable_display(self):
        for item in self.variable_treeview.get_children():
            self.variable_treeview.delete(item)
        for var_name, var_value in self.internal_variables.items():
            self.variable_treeview.insert("", tk.END, values=(f"#{var_name}", var_value))

    def act_pass(self, action_data): pass 

    def act_open_site(self, action_data):
        if not self.driver:
            logging.info("Nenhum navegador ativo. Iniciando um novo navegador...")
            browser_choice = self.browser_var.get(); is_headless = self.headless_var.get()
            try:
                if browser_choice == "Chrome":
                    chrome_options = webdriver.ChromeOptions();
                    if is_headless: chrome_options.add_argument("--headless"); chrome_options.add_argument("--window-size=1920,1080")
                    self.driver = webdriver.Chrome(options=chrome_options)
                elif browser_choice == "Firefox":
                    firefox_options = webdriver.FirefoxOptions()
                    if is_headless: firefox_options.add_argument("--headless")
                    self.driver = webdriver.Firefox(options=firefox_options)
                else: raise Exception("Navegador não suportado.")
                if not is_headless: self.driver.maximize_window()
                logging.info(f"Navegador {browser_choice} iniciado com sucesso.")
            except Exception as e:
                logging.error(f"Falha ao iniciar o navegador: {e}")
                raise
        
        url = action_data['param']
        if not url.startswith(('http://', 'https://')): url = 'http://' + url
        self.driver.get(url)

    def act_click_element(self, action_data): self._find_element(action_data).click()
    def act_click_and_wait(self, action_data):
        element = self._find_element(action_data)
        WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable(element)).click()
    def act_type_in_element(self, action_data):
        text_to_type = action_data.get('param2', '')
        element = self._find_element(action_data)
        element.clear(); element.send_keys(text_to_type)
    def act_press_enter(self, action_data): 
        self._ensure_driver()
        ActionChains(self.driver).send_keys(Keys.RETURN).perform()
    def act_select_all(self, action_data): ActionChains(self.driver).click(self._find_element(action_data)).key_down(Keys.CONTROL).send_keys('a').key_up(Keys.CONTROL).perform()
    def act_copy(self, action_data): ActionChains(self.driver).click(self._find_element(action_data)).key_down(Keys.CONTROL).send_keys('c').key_up(Keys.CONTROL).perform()
    def act_paste(self, action_data): ActionChains(self.driver).click(self._find_element(action_data)).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    def act_wait(self, action_data):
        wait_time = action_data.get('param','0')
        time.sleep(float(wait_time))
        
    def act_wait_for_human(self, action_data):
        self.dialog_event.clear()
        dialog_call = partial(messagebox.askquestion, "Verificação Humana Necessária", "O processo foi verificado. Deseja continuar com as ações?", icon='question')
        self.root.after(0, self._execute_dialog_on_main_thread, dialog_call)
        self.dialog_event.wait()
        response = self.dialog_result
        if response == 'no': raise UserCancelledException("Usuário cancelou a automação.")
        
    def act_save_clipboard_to_variable(self, action_data):
        variable_name = action_data.get('param')
        if not variable_name:
            logging.warning("Ação 'Salvar Área de Transferência em Variável' chamada sem nome de variável.")
            return

        try:
            clipboard_content = self.root.clipboard_get()
            
            # Lida com variáveis internas (#) e de dados ($)
            if variable_name.startswith('#'):
                var_key = variable_name[1:]
                self.internal_variables[var_key] = clipboard_content
                logging.info(f"Conteúdo da área de transferência salvo na variável interna '{variable_name}'.")
                self.root.after(0, self._update_variable_display)

            elif variable_name.startswith('$'):
                if not self.imported_data: 
                    logging.warning("Tentativa de salvar na variável de dados, mas nenhum dado foi importado.")
                    return
                variable_pattern = r'\$([a-zA-Z_][a-zA-Z0-9_]*)(\d+)'
                match = re.match(variable_pattern, variable_name)
                if not match: raise ValueError("Formato da variável de dados inválido.")
                
                col_name = match.group(1)
                row_num = int(match.group(2))
                col_index = self.imported_data["headers"].index(col_name)
                row_index = row_num - 1

                while len(self.imported_data["rows"]) <= row_index: self.imported_data["rows"].append([''] * len(self.imported_data["headers"]))
                while len(self.imported_data["rows"][row_index]) <= col_index: self.imported_data["rows"][row_index].append('')
                
                self.imported_data["rows"][row_index][col_index] = clipboard_content
                logging.info(f"Conteúdo '{clipboard_content[:20]}...' salvo na variável de dados {variable_name}.")
            else:
                logging.warning(f"Formato de variável desconhecido para: {variable_name}. Use '#' ou '$'.")
        
        except (ValueError, IndexError) as e: 
            logging.error(f"Não foi possível salvar na variável {variable_name}. Erro: {e}")
        except Exception as e: 
            logging.error(f"Erro inesperado ao salvar na variável: {e}")

    # --- NOVO ---
    def act_save_variable_to_clipboard(self, action_data):
        """Resolve uma variável e salva seu conteúdo na área de transferência."""
        variable_name = action_data.get('param')
        if not variable_name:
            logging.warning("Ação 'Salvar Variável na Área de Transferência' chamada sem nome de variável.")
            return

        try:
            resolved_value = str(self._resolve_variable(variable_name))
            
            self.root.clipboard_clear()
            self.root.clipboard_append(resolved_value)
            
            logging.info(f"Valor da variável '{variable_name}' ('{resolved_value[:50]}...') salvo na área de transferência.")
        except Exception as e:
            logging.error(f"Erro ao salvar variável '{variable_name}' na área de transferência: {e}")
            raise

    def act_switch_tab(self, action_data):
        self._ensure_driver()
        tab_index = int(action_data.get('param', '0'))
        if len(self.driver.window_handles) > tab_index:
            self.driver.switch_to.window(self.driver.window_handles[tab_index])
            logging.info(f"Trocado para a aba {tab_index}.")
        else: logging.warning(f"Aba {tab_index} não encontrada.")

    def act_close_tab(self, action_data):
        self._ensure_driver()
        if len(self.driver.window_handles) > 1:
            self.driver.close()
            self.driver.switch_to.window(self.driver.window_handles[-1])
            logging.info("Aba fechada.")
        else: logging.warning("Não é possível fechar a última aba.")

    def act_upload_file(self, action_data):
        file_path = action_data.get("param2", "")
        if not os.path.isabs(file_path): file_path = os.path.join(os.getcwd(), file_path)
        if os.path.exists(file_path):
            element = self._find_element(action_data)
            element.send_keys(file_path)
            logging.info(f"Arquivo '{file_path}' enviado para o elemento.")
        else:
            logging.error(f"Arquivo de upload não encontrado em: {file_path}")
            raise FileNotFoundError(f"Arquivo de upload não encontrado: {file_path}")
            
    def act_create_internal_variable(self, action_data):
        var_name = action_data.get('param')
        initial_value = action_data.get('param2', '')
        if not var_name:
            logging.warning("Ação 'Criar Variável Interna' chamada sem nome de variável.")
            return
        resolved_value = self._resolve_variable(initial_value)
        self.internal_variables[var_name] = resolved_value
        logging.info(f"Variável interna '{var_name}' criada/atualizada com o valor '{resolved_value}'.")
        self.root.after(0, self._update_variable_display)

    def act_prompt_user_input(self, action_data):
        prompt_message = action_data.get('param', 'Por favor, insira um valor:')
        var_name = action_data.get('param2')
        if not var_name:
            logging.warning("Ação 'Pedir Input do Usuário' chamada sem um nome de variável para salvar.")
            return
        
        resolved_prompt = self._resolve_variable(prompt_message)
        
        self.dialog_event.clear()
        dialog_call = partial(simpledialog.askstring, "Input do Usuário Necessário", resolved_prompt)
        self.root.after(0, self._execute_dialog_on_main_thread, dialog_call)
        self.dialog_event.wait()
        user_input = self.dialog_result
        
        self.internal_variables[var_name] = user_input if user_input is not None else ""
        logging.info(f"Input do usuário '{self.internal_variables[var_name]}' salvo na variável interna '{var_name}'.")
        self.root.after(0, self._update_variable_display)

    def act_wait_for_browser_close(self, action_data):
        self._ensure_driver()
        logging.info("Aguardando o usuário fechar o navegador para continuar...")
        while True:
            try:
                if not self.driver.window_handles: break
                time.sleep(1)
            except (InvalidSessionIdException, ProtocolError, ConnectionRefusedError):
                logging.info("Navegador fechado pelo usuário. Continuando a execução.")
                break
            except Exception as e:
                logging.error(f"Erro inesperado ao aguardar fechamento do navegador: {e}")
                break
        self.driver = None 

    def act_extract_text(self, action_data):
        var_name = action_data.get('param2')
        if not var_name:
            logging.warning("Ação 'Extrair Texto' chamada sem nome de variável para salvar.")
            return
        element = self._find_element(action_data)
        extracted_text = element.text
        self.internal_variables[var_name] = extracted_text
        logging.info(f"Texto '{extracted_text[:30]}...' extraído e salvo na variável '{var_name}'.")
        self.root.after(0, self._update_variable_display)

    def act_extract_attribute(self, action_data):
        attribute_name = action_data.get('param2')
        var_name = action_data.get('param3')
        if not attribute_name or not var_name:
            logging.warning("Ação 'Extrair Atributo' chamada sem nome do atributo ou da variável.")
            return
        element = self._find_element(action_data)
        attribute_value = element.get_attribute(attribute_name)
        self.internal_variables[var_name] = attribute_value
        logging.info(f"Atributo '{attribute_name}' com valor '{attribute_value}' salvo na variável '{var_name}'.")
        self.root.after(0, self._update_variable_display)

    # CÓDIGO A SER ADICIONADO
    def act_extract_table_to_csv(self, action_data):
        """
        Encontra uma tabela na página, extrai seus dados (cabeçalhos e linhas)
        e salva em um arquivo CSV no diretório base da aplicação.
        """
        csv_filename = action_data.get('param2')
        if not csv_filename:
            raise ValueError("O nome do arquivo CSV de destino não foi especificado na ação.")

        if not csv_filename.lower().endswith('.csv'):
            csv_filename += '.csv'

        output_path = os.path.join(self.base_path, csv_filename)
        logging.info(f"Iniciando extração de tabela para o arquivo: {output_path}")

        # Encontra o elemento da tabela usando a lógica de seletor já existente
        table_element = self._find_element(action_data)

        data = []

        # Procura por cabeçalhos (th) dentro da tabela
        headers = table_element.find_elements(By.TAG_NAME, 'th')
        if headers:
            header_row = [header.text.strip() for header in headers]
            data.append(header_row)

        # Procura por todas as linhas (tr) da tabela
        rows = table_element.find_elements(By.TAG_NAME, 'tr')
        for row in rows:
            # Dentro de cada linha, procura por todas as células (td)
            cells = row.find_elements(By.TAG_NAME, 'td')
            if cells:  # Apenas processa linhas que contêm células de dados
                row_data = [cell.text.strip().replace('\n', ' ') for cell in cells]
                data.append(row_data)

        if not data:
            logging.warning(f"A tabela encontrada com o seletor '{action_data.get('param')}' parece estar vazia. Nenhum arquivo foi gerado.")
            return

        # Escreve os dados coletados no arquivo CSV
        try:
            with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
                # Usamos ';' como delimitador, comum no Brasil para abrir corretamente no Excel
                writer = csv.writer(csvfile, delimiter=';')
                writer.writerows(data)
            logging.info(f"Tabela extraída com sucesso para '{output_path}'. {len(data)} linhas salvas.")
            # A notificação na tela precisa ser chamada na thread principal do Tkinter
            self.root.after(0, messagebox.showinfo, "Extração Concluída", f"A tabela foi salva com sucesso em:\n{output_path}")
        except Exception as e:
            error_msg = f"Falha ao escrever o arquivo CSV em '{output_path}'. Erro: {e}"
            logging.error(error_msg)
            raise IOError(error_msg)

    def act_manipulate_variable(self, action_data):
        dest_var = action_data.get('param')
        expression = action_data.get('param2', '')
        if not dest_var or not expression:
            logging.warning("Ação 'Manipular Variável' chamada sem variável de destino ou expressão.")
            return
        
        try:
            resolved_expression = self._resolve_variable(expression)
            result = None
            if '+' in resolved_expression:
                parts = [p.strip() for p in resolved_expression.split('+', 1)]
                try: result = float(parts[0]) + float(parts[1])
                except ValueError: result = str(parts[0]) + str(parts[1])
            elif '-' in resolved_expression:
                parts = [p.strip() for p in resolved_expression.split('-', 1)]
                result = float(parts[0]) - float(parts[1])
            elif '*' in resolved_expression:
                parts = [p.strip() for p in resolved_expression.split('*', 1)]
                result = float(parts[0]) * float(parts[1])
            elif '/' in resolved_expression:
                parts = [p.strip() for p in resolved_expression.split('/', 1)]
                result = float(parts[0]) / float(parts[1])
            else:
                result = resolved_expression
            
            self.internal_variables[dest_var] = result
            logging.info(f"Variável '{dest_var}' atualizada para '{result}' pela expressão '{expression}'.")
            self.root.after(0, self._update_variable_display)
        except Exception as e:
            logging.error(f"Erro ao manipular variável '{dest_var}' com expressão '{expression}': {e}")
            raise ValueError(f"Erro na expressão de 'Manipular Variável': {e}")
            
    def act_hover_element(self, action_data):
        element = self._find_element(action_data)
        ActionChains(self.driver).move_to_element(element).perform()
        logging.info("Ação de 'Hover' (mover mouse) executada.")

    def act_double_click(self, action_data):
        element = self._find_element(action_data)
        ActionChains(self.driver).double_click(element).perform()
        logging.info("Ação de 'Duplo Clique' executada.")
        
    def act_scroll_page(self, action_data):
        self._ensure_driver()
        direction = action_data.get('param', 'fim').lower()
        if direction == 'fim':
            self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            logging.info("Página rolada para o fim.")
        elif direction == 'topo':
            self.driver.execute_script("window.scrollTo(0, 0);")
            logging.info("Página rolada para o topo.")
        else:
            try:
                pixels = int(direction)
                self.driver.execute_script(f"window.scrollBy(0, {pixels});")
                logging.info(f"Página rolada por {pixels} pixels.")
            except ValueError:
                logging.warning(f"Direção de rolagem inválida: '{direction}'. Use 'topo', 'fim' ou um número.")

    def act_scroll_to_element(self, action_data):
        element = self._find_element(action_data)
        self.driver.execute_script("arguments[0].scrollIntoView(true);", element)
        logging.info("Página rolada até o elemento.")

    def act_take_screenshot(self, action_data):
        self._ensure_driver()
        filename = action_data.get('param')
        if not filename:
            timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            filename = f"screenshot_{timestamp}.png"
        
        filepath = os.path.join(self.base_path, filename)
        try:
            self.driver.save_screenshot(filepath)
            logging.info(f"Screenshot salvo em: {filepath}")
        except Exception as e:
            logging.error(f"Falha ao salvar screenshot em '{filepath}': {e}")
            raise IOError(f"Não foi possível salvar a captura de tela: {e}")

    def act_execute_js(self, action_data):
        self._ensure_driver()
        script = action_data.get('param', '')
        if script:
            try:
                self.driver.execute_script(script)
                logging.info(f"Script JavaScript executado: {script[:50]}...")
            except Exception as e:
                logging.error(f"Falha ao executar JavaScript: {e}")
                raise
    
    def act_refresh_page(self, action_data):
        self._ensure_driver()
        self.driver.refresh()
        logging.info("Página atualizada (F5).")

    def act_navigate_back(self, action_data):
        self._ensure_driver()
        self.driver.back()
        logging.info("Navegou para a página anterior.")

    def act_navigate_forward(self, action_data):
        self._ensure_driver()
        self.driver.forward()
        logging.info("Navegou para a página seguinte.")
    
    def act_click_at_xy(self, action_data):
        try:
            x = int(action_data.get('param'))
            y = int(action_data.get('param2'))
            logging.info(f"Clicando nas coordenadas de tela (X={x}, Y={y}).")
            pyautogui.click(x=x, y=y)
        except (ValueError, TypeError):
            msg = f"Coordenadas inválidas para 'Clicar em Coordenadas': X='{action_data.get('param')}', Y='{action_data.get('param2')}'. Devem ser números."
            logging.error(msg)
            raise ValueError(msg)
        except Exception as e:
            logging.error(f"Erro ao clicar em coordenadas com pyautogui: {e}")
            raise

    def act_click_at_xy_and_type(self, action_data):
        try:
            x = int(action_data.get('param'))
            y = int(action_data.get('param2'))
            text_to_type = action_data.get('param3', '')
            logging.info(f"Clicando em (X={x}, Y={y}) e digitando '{text_to_type}'.")
            pyautogui.click(x=x, y=y)
            pyautogui.write(text_to_type)
        except (ValueError, TypeError):
            msg = f"Coordenadas inválidas para 'Clicar e Escrever': X='{action_data.get('param')}', Y='{action_data.get('param2')}'. Devem ser números."
            logging.error(msg)
            raise ValueError(msg)
        except Exception as e:
            logging.error(f"Erro ao clicar em coordenadas e escrever com pyautogui: {e}")
            raise

    def act_press_multiple_keys(self, action_data):
        keys_str = action_data.get('param', '').lower().strip()
        if not keys_str:
            logging.warning("Nenhuma tecla especificada para 'Pressionar Múltiplas Teclas'.")
            return

        logging.info(f"Pressionando as teclas de atalho: '{keys_str}'")
        
        try:
            keys_to_press = [p.strip() for p in keys_str.split('+')]
            pyautogui.hotkey(*keys_to_press)
        except Exception as e:
            logging.error(f"Erro ao pressionar teclas de atalho com pyautogui: {e}")
            raise

    def act_type_globally(self, action_data):
        """
        Usa pyautogui para digitar um texto onde o cursor estiver focado na tela.
        Independente do navegador.
        """
        text_to_type = action_data.get('param')

        # Verifica se o parâmetro foi fornecido. Uma string vazia é um valor válido.
        if text_to_type is None:
            logging.warning("Ação 'Digitar Texto (Global)' chamada sem texto para digitar.")
            return

        try:
            # Resolve qualquer variável no texto (ex: #minha_var ou $Coluna1) antes de digitar
            resolved_text = self._resolve_variable(text_to_type)
            logging.info(f"Digitando texto globalmente: '{resolved_text}'")

            # O intervalo adiciona uma pequena pausa entre as teclas, tornando a digitação mais confiável
            pyautogui.write(resolved_text, interval=0.05)

        except Exception as e:
            logging.error(f"Erro ao digitar globalmente com pyautogui: {e}")
            raise

    def on_drag_start(self, event):
        widget = event.widget; self.drag_start_index = widget.nearest(event.y)
    def on_drag_motion(self, event):
        if self.drag_start_index is None: return
        widget = event.widget; new_index = widget.nearest(event.y)
        if new_index != self.listbox.curselection():
            self.listbox.selection_clear(0, tk.END); self.listbox.selection_set(new_index); self.listbox.activate(new_index)
    def on_drop(self, event):
        if self.drag_start_index is None or not self.active_profile: return
        drop_index = event.widget.nearest(event.y)
        if drop_index == self.drag_start_index: self.drag_start_index = None; return
        actions = self.profiles_data["profiles"][self.active_profile]["actions"]
        item_to_move = actions.pop(self.drag_start_index)
        actions.insert(drop_index, item_to_move)
        logging.info(f"Ação movida de {self.drag_start_index + 1} para {drop_index + 1}.")
        self.drag_start_index = None
        self.update_listbox()


class UserCancelledException(Exception): pass

if __name__ == "__main__":
    main_window = tk.Tk()
    app = AutomationGUI(main_window)
    main_window.mainloop()