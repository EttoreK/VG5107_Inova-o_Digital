# Módulos necessários
import customtkinter as ctk # aplicativo
import os # controle de arquivos do sistema

# imports dos arquivos criados para cada problema
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

        # Configurações das telas do aplicativo
        self.container = ctk.CTkFrame(self) 
        self.container.pack(fill="both", expand=True) # preenche toda a janela om a tela

        self.tela = Resolucao(parent=self.container)
        self.tela.pack(fill="both", expand=True)
        
# Classe/Tela para cchamar o script que resolve e conversar com o usuário
class Resolucao(ctk.CTkFrame):
    def __init__(self, paren):
        super().__init__(parent)
        self.problema_ativo = None
        self.tipo_arquivo = ""
        self.arqDir = ""

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
            command=lambda: self.rodar(problema3, "pdf"))
        self.btnP1.pack(pady=10)

        self.btnP2 = ctk.CTkButton(
            self, text="Problema 2", 
            # lambda para executar só depois de clicado
            command=lambda: self.rodar(problema3, "excel"))
        self.btnP2.pack(pady=10)
        
        self.btnP3 = ctk.CTkButton(
            self, text="Problema 3", 
            # lambda para executar só depois de clicado
            command=lambda: self.rodar(problema3, "excel"))
        self.btnP3.pack(pady=10)

        # Título escrito dentro da tela
        self.titulo = ctk.CTkLabel(
            self, text="Resolução do Problema", 
            font=("Roboto", 16, "bold")
        )
        self.titulo.pack(pady=(42, 10))

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
    
        self.logar("Aguardando arquivo...")
    
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
        
        # Define o tipo de arquivo para ser escolhido
        if self.tipo_arquivo == "excel":
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
        self.update()

        # Verifica se o módulo foi importado, existe e pode ser chamado
        try:
            modulo = self.problema_ativo
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

# Classe main, chama o app
if __name__ == "__main__":
    app = App()
    app.mainloop()