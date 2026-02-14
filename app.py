# Módulos necessários
import customtkinter as ctk # aplicativo
import os # controle de arquivos do sistema

# imports dos arquivos criados para cada problema
import problema1 # pdf
import problema2 # tratamento
import problema3 # melt

# Configurações de tema do aplicativo
ctk.set_appearance_mode("System") 
ctk.set_default_color_theme("dark-blue")  

# Classe principal faz-tudo
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        # Configurações da janela do app
        self.title("") # título da janela
        self.geometry("256x256") # tamanho da janela
        self.resizable(False, False) # se é redimensionávvel

        # Atributos para determinar as ações de algumas funções
        self.arqTipo = ""
        self.problema = None

        # Configurações das telas do aplicativo
        self.container = ctk.CTkFrame(self) 
        self.container.pack(fill="both", expand=True) # preenche toda a janela om a tela
        self.container.grid_rowconfigure(0, weight=1) # 100% de preenchimento
        self.container.grid_columnconfigure(0, weight=1) # 100% de preenchimento

        # Criação das telas com base nas classes
        self.telas = {}
        for t in (Menu, Resolusao):
            nomeTela = t.__name__
            tela = t(parent=self.container, controller=self) # inicia as classes
            self.telas[nomeTela] = tela
            tela.grid(row=0, column=0, sticky="nsew") #1 linha 1 coluna norte-sul-leste-oeste

        self.escolherTela("Menu") # tela inicial Menu

    # Altera a tela exibida
    def escolherTela(self, nomeTela):
        """Exibe a tela escolhida"""
        tela = self.telas[nomeTela]
        
        # Tratamento se for de Resolucao para resolucao, só reinicia a tela
        if hasattr(tela, "mostrarTela"):
            tela.mostrarTela()

        tela.tkraise() # exibe a tela

    # Seleção entre os 3 problemas possíveis
    def escolherProblema(self, problema, tipo):
        """Guarda a escolha e troca para tela Resolusao"""
        self.problema = problema
        self.arqTipo = tipo
        self.escolherTela("Resolusao")

# Classe/Tela para escolher o problema
class Menu(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        # Título escrito dentro da tela
        self.titulo = ctk.CTkLabel(
            self, text="Escolha o Problema", 
            font=("Roboto", 16, "bold")
        )
        self.titulo.pack(pady=(20, 30))

        # Botões de escolha
        self.btnP1 = ctk.CTkButton(
            self, text="Problema 1", 
            # lambda para executar só depois de clicado
            command=lambda: controller.escolherProblema(problema1, "pdf"))
        self.btnP1.pack(pady=10)

        self.btnP2 = ctk.CTkButton(
            self, text="Problema 2", 
            # lambda para executar só depois de clicado
            command=lambda: controller.escolherProblema(problema2, "excel"))
        self.btnP2.pack(pady=10)
        
        self.btnP3 = ctk.CTkButton(
            self, text="Problema 3", 
            # lambda para executar só depois de clicado
            command=lambda: controller.escolherProblema(problema3, "excel"))
        self.btnP3.pack(pady=10)

# Classe/Tela para cchamar o script que resolve e conversar com o usuário
class Resolusao(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.arqDir = ""

        # Retorna ao Menu
        self.btnVoltar = ctk.CTkButton(
            self, text="Voltar", 
            width=60, 
            fg_color="gray", 
            command=self.vaoltarMenu
        )
        self.btnVoltar.place(x=8, y=8)

        # Título escrito dentro da tela
        self.titulo = ctk.CTkLabel(
            self, text="Resolução do Problema", 
            font=("Roboto", 16, "bold")
        )
        self.titulo.pack(pady=(42, 10))

        # Botão para escolher o arquivo
        self.btnArq = ctk.CTkButton(
            self, text="Escolher Arquivo", 
            command=self.escolherArquivo
        )
        self.btnArq.pack(pady=10)

        # Acompanhamento do arquivo selecionado
        self.labelArq = ctk.CTkLabel(
            self, text="Nenhum arquivo selecionado", 
            text_color="gray"
        )
        self.labelArq.pack(pady=5)

        # Acompanhamento da resolução e controle para o usuário
        self.txtboxLog = ctk.CTkTextbox(self, width=240, height=80)
        self.txtboxLog.pack(pady=10)
        self.txtboxLog.configure(state="disabled") # não permite ninguém alem do sistema escrever
        
        # Botão para copíar o log
        self.btnCopiar = ctk.CTkButton(
            self, text="⎘", 
            width=10, height=10, 
            fg_color="gray", hover_color="black",
            command=self.copiarLog
        )
        self.btnCopiar.place(x=228, y=160)

        self.logar("Aguardando arquivo...")
    
    # (Re)Exibe a tela de resolução para limpar o log do processo anterior
    def mostrarTela(self):
        """Limpa o log e o arquivo escolhido ao reentrar na tela"""
        self.arqDir = ""
        self.txtboxLog.delete("0.0", "end") # apaga tudo
        self.labelArq.configure(text="Nenhum arquivo escolhido", text_color="gray")
        self.btnArq.configure(state="normal")
        self.btnVoltar.configure(state="normal")

        # Muda o título pada cada problema
        self.titulo.configure(
            text=(
                "Resolução do %s" % self.controller.problema.__name__.capitalize()
            )
        )
        
        # Atualiza o texto dependendo do tipo escolhido
        tipo = self.controller.arqTipo.upper()
        self.btnArq.configure(text=("Escolher  %s" % tipo))
        self.logar("Aguardando arquivo %s..." % tipo)
    
    # Escreve para o usuário ver em qual parte o processo está
    def logar(self, mensagem):
        """Adiciona no final da caixa de texto a mensagem"""
        self.txtboxLog.configure(state="normal")
        self.txtboxLog.insert("end", ("> %s\n" % mensagem))
        self.txtboxLog.configure(state="disabled")
        self.txtboxLog.see("end") # mostra as últimas linhas
        self.update_idletasks() # redesenha a tela 

    # Usuário escolhe o arquivo de entrada e local de saída
    def escolherArquivo(self):
        """Pega o arquivo e local de salvar"""
        tipo = self.controller.arqTipo
        arquivo = ""
        
        # Define o tipo de arquivo para ser escolhido
        if tipo == "excel":
            arquivo = ctk.filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        else:
            arquivo = ctk.filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])

        # Se o arquivvo não for vazio 
        if arquivo:
            self.arqDir = arquivo
            self.labelArq.configure(text=("...%s" % os.path.basename(arquivo)), text_color="gray")
            self.logar("Selecionado: %s" % os.path.basename(arquivo))

                # Padrões de saída, nome sempre igual e saída sempre excel
                extensao = ".xlsx"
                nomePadrao = "Relatório_" + os.path.splitext(os.path.basename(self.arqDir))[0]

                # Pergunta aonde salvar
                arquivoSaida = ctk.filedialog.asksaveasfilename(
                    defaultextension=extensao,
                    initialfile=nomePadrao,
                    filetypes=[("Excel", "*.xlsx")],
                    title="Escolha onde salvar o relatório"
                )

            # Se tem saída, roda a resolução
            if arquivoSaida:
                self.rodar(arquivoSaida)
            else:
                self.logar("Operação cancelada pelo usuário.")
    
    # Chama o arquivo que resolve o problema
    def rodar(self, arquivoSaida):
        """Executa uma soloução para o problema escolhido"""
        # Evita que o código quebre ou trave
        self.btnArq.configure(state="disabled")
        self.btnVoltar.configure(state="disabled")
        self.update()

        # Verifica se o módulo foi importado, existe e pode ser chamado
        try:
            modulo = self.controller.problema
            if not modulo:
                self.logar("Erro: Módulo não encontrado.")
                return

            self.logar("Iniciando...")

            # Controle de erros
            sucesso, mensagem = modulo.executar(self.arqDir, arquivoSaida, self.logar)

            if not sucesso:
                self.logar("ERRO: %s" % mensagem)
        
        # Segurança para qualquer coisa
        except Exception as e:
            self.logar("ERRO CRÍTICO: %s" % str(e))
        
        # Libera para tentar novvamente 
        finally:
            self.btnArq.configure(state="normal")
            self.btnVoltar.configure(state="normal")

# Classe main, chama o app
if __name__ == "__main__":
    app = App()
    app.mainloop()