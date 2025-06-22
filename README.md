WebStride Automatizador

WebStride é uma poderosa ferramenta de automação web com interface gráfica (GUI), construída em Python usando Tkinter e Selenium. Ela capacita usuários, mesmo sem conhecimento profundo de programação, a criar, gerenciar e executar macros complexas para automatizar tarefas repetitivas em qualquer website.

O sistema combina a precisão do Selenium para interagir com elementos da web, a flexibilidade do PyAutoGUI para controle global do desktop e uma lógica de programação visual para criar fluxos de trabalho robustos.

📸 Vitrine da Aplicação
(Recomendação: Substitua o texto abaixo por um screenshot da interface principal do WebStride em execução)

+-------------------------------------------------------------------------+
| WebStride Automatizador                        [⚙] [?] [_] [X]          |
+-------------------------------------------------------------------------+
| Gerenciar Perfis: [Perfil Padrão ▼] [Novo] [Renomear] [Excluir]           |
| [Importar Dados] C:\Data\users.csv                                      |
+-------------------------------------------------------------------------+
|  Navegador:     Ação:                                                   |
|  (•) Chrome     [Clicar em Elemento                                ▼]   |
|  ( ) Firefox    Nome/Descrição: [Login no sistema]                        |
|  [ ] Headless   Tipo de Seletor: [CSS Selector ▼] Seletor: [#loginButton] |
|                 ...                                                     |
|                                                     [Adicionar Ação]    |
+-------------------------------------------------------------------------+
| ▼ Sequência de Ações do Perfil: 'Meu Perfil'                            |
| | 1. [Acessar site] Abrir Site: https://meusite.com                     |
| | 2. [Preencher user] Escrever em Campo: [ID] user | Valor: $Usuario1  |
| | 3. [Preencher senha] Escrever em Campo: [ID] pass | Valor: $Senha1   |
| | 4. Clicar em Elemento: [CSS Selector] #loginButton                    |
| |                                                                       |
+-------------------------------------------------------------------------+
| [Editar Ação] [->] [<-] [Remover]            [Salvar] [■ Parar] [▶ Executar] |
+-------------------------------------------------------------------------+
✨ Funcionalidades Principais
Interface Visual Intuitiva: Crie automações complexas arrastando e soltando (futuro), editando e organizando ações em uma lista sequencial, sem escrever uma linha de código.
Gerenciamento de Perfis: Salve diferentes sequências de automação como "Perfis". Cada perfil pode ter sua própria macro, navegador de preferência e arquivo de dados associado.
Sistema de Ações Abrangente: Uma vasta gama de ações categorizadas para cobrir todas as necessidades de automação:
Navegador e Página: Abrir sites, atualizar, navegar, gerenciar abas, executar JavaScript.
Interação com Elementos (Selenium): Clicar, escrever, fazer upload de arquivos, mover o mouse (hover), rolar até elementos, etc.
Ações Globais (PyAutoGUI): Clicar em coordenadas específicas da tela, digitar texto globalmente e pressionar atalhos de teclado (ex: Ctrl+S).
Extração de Dados: Extraia texto de elementos, valores de atributos e tabelas HTML inteiras diretamente para arquivos .csv.
Controle de Fluxo: Implemente lógica em suas macros com laços de repetição (Loop, Loop Fixo) e condicionais (Se, Senão Se, Senão).
Variáveis e Área de Transferência: Crie variáveis internas, peça inputs do usuário durante a execução, e manipule a área de transferência do sistema.
Automação Guiada por Dados (Data-Driven):
Importe dados de arquivos .txt ou .csv (com delimitador ;).
Utilize um sistema de variáveis dinâmicas (ex: $Coluna1) para iterar sobre os dados importados, permitindo preencher formulários, realizar buscas ou extrair informações em massa.
Variáveis Internas: Crie e manipule variáveis internas (ex: #contador) para armazenar e reutilizar informações durante a execução da macro.
Testador de Seletores em Tempo Real: Abra um "Navegador de Teste" e use a função "Testar" para verificar e destacar em tempo real se seus seletores (CSS, XPath, ID, etc.) estão corretos antes de executar a automação completa.
Controle de Execução: Inicie, pare o navegador a qualquer momento e acompanhe a execução que destaca a ação atual na lista.
Human-in-the-Loop: Pause a automação para solicitar verificação humana ou para que o usuário insira um dado específico através de uma caixa de diálogo.
Logging Detalhado: A aplicação gera automaticamente arquivos log.txt e log_erro.txt no diretório C:\WebStride, facilitando a depuração de problemas.
Persistência: Todas as configurações e perfis são salvos em um arquivo database.json, garantindo que seu trabalho nunca seja perdido.
⚙️ Como Funciona
O WebStride opera sobre uma pilha de execução que interpreta a lista de ações que você construiu.

Criação do Perfil: Você cria um perfil para sua tarefa.
Construção da Macro: Você adiciona ações da lista suspensa, configurando parâmetros como URLs, seletores de elementos, texto para digitar, etc.
Lógica de Controle: Você pode recuar/indentar ações para aninhá-las dentro de blocos de controle de fluxo, como Iniciar Loop ou Se (condição).
Associação de Dados (Opcional): Você pode importar um arquivo .csv/.txt. As ações dentro de um Iniciar Loop irão iterar sobre as linhas deste arquivo, substituindo variáveis como $Coluna{N} pelos dados correspondentes da linha atual.
Execução: Ao clicar em "Executar", um thread separado inicia o driver do Selenium (Chrome ou Firefox, visível ou headless) e começa a processar a lista de ações, uma por uma, resolvendo variáveis e executando as interações na página web.
🚀 Instalação e Configuração
Para executar o WebStride, você precisará do Python 3 e de algumas bibliotecas.

1. Pré-requisitos
Python 3.7+
Um navegador web (Google Chrome ou Mozilla Firefox).
2. Clonar o Repositório
Bash

git clone https://github.com/seu-usuario/WebStride.git
cd WebStride
3. Instalar Dependências
O projeto utiliza algumas bibliotecas Python que podem ser instaladas via pip.

Bash

pip install selenium pyautogui
4. Configurar o WebDriver
O Selenium precisa de um "motor" (WebDriver) para controlar o navegador.

Google Chrome: Baixe o ChromeDriver correspondente à sua versão do Chrome aqui.
Mozilla Firefox: Baixe o GeckoDriver aqui.
Importante: Após o download, descompacte o arquivo e coloque o executável (chromedriver.exe ou geckodriver.exe) em um local que esteja no PATH do seu sistema (como C:\Windows) ou na mesma pasta onde o script WebStride.py está localizado. Isso garante que o Selenium consiga encontrá-lo.

📖 Guia de Uso Rápido
Execute o Script:
Bash

python WebStride.py
Crie seu Primeiro Perfil:
A aplicação iniciará com um "Perfil Padrão". Você pode renomeá-lo ou criar um novo clicando em "Novo".
Adicione Ações:
Ação "Abrir Site": No campo "Parâmetro", digite google.com. Clique em "Adicionar Ação".
Ação "Escrever em Campo":
Tipo de Seletor: Atributo name
Seletor: q
Valor: Automação com Selenium
Clique em "Adicionar Ação".
Ação "Pressionar Enter": Adicione esta ação.
Execute a Automação:
Selecione o navegador de sua preferência (Chrome/Firefox).
Clique no botão verde ▶ Executar.
Observe a Mágica: Uma nova janela do navegador será aberta e executará as ações que você definiu!
🗂️ Estrutura de Arquivos Gerada
Ao ser executado pela primeira vez, o WebStride criará a seguinte estrutura de diretórios em sua máquina:

C:\
└── WebStride\
    ├── Database\
    │   └── database.json  # Arquivo principal que armazena todos os perfis e ações
    ├── log.txt            # Log de eventos e execução bem-sucedida
    └── log_erro.txt       # Log de erros e exceções
Qualquer arquivo gerado pela automação (como screenshots ou tabelas CSV) também será salvo dentro de C:\WebStride.

🤝 Contribuições
Contribuições são muito bem-vindas! Se você tem ideias para novas funcionalidades, melhorias na interface ou encontrou um bug, sinta-se à vontade para:

Fazer um fork do projeto.
Criar uma nova branch (git checkout -b feature/NovaFuncionalidade).
Fazer o commit de suas alterações (git commit -am 'Adiciona NovaFuncionalidade').
Fazer o push para a branch (git push origin feature/NovaFuncionalidade).
Abrir um Pull Request.
📜 Licença
Este projeto está sob a licença MIT. Veja o arquivo LICENSE.md para mais detalhes.







