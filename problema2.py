# Módulos necessários
import pandas as pd
import xlwings as xw
import os

# Lê o workbook do excel
def lerWorkbook(app, arquivo):
    """Captura o workbook do arquivo excel"""
    arquivo = arquivo.replace('\\ ', ' ').strip() # garante que o caminho é legível

    # Se o arquivo não existe, falha
    if not os.path.exists(arquivo):
        print("Falha ao ler arquivo, cancelando operação.")
        return False, False
    
    # Tenta renomear o arquivo para evr se ele está em estado editável
    try:
        os.rename(arquivo, arquivo)
    except:
        print("O arquivo está aberto em outro processo, cancelando operação.")
        return False, False
    
    workbook = app.books.open(arquivo) # abre para edição

    return workbook, arquivo

# Lê as planilhas do workbook
def lerPlanilhas(workbook):
    """Lê a as planilhas do arquivvo Excel"""
    # Se o work book não foi aberto corretamente, falha
    if not workbook:
        print("Erro ao acessar o arquivo, cancelando operação.")
        return False, False, False
    
    # Nomes padrão
    planilha_valores = "Valores_orientadores"
    planilha_risco = "Avaliacao_Risco_Case"

    # Extrai os nomes das planilhas
    planilhas = workbook.sheets
    planilhas_arquivo = {}
    for planilha in planilhas:
        indice = planilha.index
        nome = planilha.name
        planilhas_arquivo[indice] = nome
        print("%s\t%s" % (indice, nome))

    # Se os nomes padão não estiverem na planilha, falha
    if ((planilha_valores not in planilhas_arquivo.keys()) and (planilha_valores not in planilhas_arquivo.values()) or 
        ((planilha_risco not in planilhas_arquivo.keys()) and (planilha_risco not in planilhas_arquivo.values()))):
        print("Nome ou número da planilha não encontrados, cancelando operação.")
        return False, False, False

    # Extrai as planilhas individualmente
    ws1 = workbook.sheets[planilha_valores]
    ws2 = workbook.sheets[planilha_risco]

    return ws1, ws2, planilhas_arquivo

# Cria uma nova planilha para inserir os valores
def criarPlanilha(workbook, planilhas_existe):
    """Cria uma nova planilha para inserir valores"""
    # Se o workbook não existir, ou as planilhas extraidas, falha 
    if not workbook or not planilhas_existe:
        print("Erro ao acessar o arquivo, cancelando operação.")
        return False
    
    # Cria nova planilha com nome padrão
    planilha_nome = "Avaliacao_Risco"
    ws = workbook.sheets.add(planilha_nome)

    return ws

# Extrai os dados da planilha em foram de dataframe
def coletaDados(planilha_base, arquivo, pula):
    """Coleta os dado da planilha no arquivo e pula as linhas indesejadas"""
    # Falha se não existir
    if not planilha_base or not arquivo:
        print("Erro ao ler planilha, cancelando operação.")
        return False
    
    # Planilha para dataframe
    df = pd.read_excel(
        arquivo, sheet_name=planilha_base.name, 
        index_col=False, skiprows=pula
    )

    # Se não conseguiu extrair, falha
    if df.empty:
        print("Erro ao coletar dados, cancelando operação.")
        return False
    
    return df

# Transforma o dataframe pora ter a cara certa
def transformarDados(dados_valores, dados_risco):
    """Transforma os dados para terem o farmato certo"""
    # Falha se algo faltar
    if dados_valores.empty or dados_risco.empty:
        print("Erro ao transformar dados, cancelando operação.")
        return False

    if "Parâmetro" not in dados_valores.columns:
        dados_valores = dados_valores.rename(columns={dados_valores.columns[1]: "Parâmetro"})
    
    # dados_valores.iloc[0:, 1] = dados_valores.iloc[:, 1].astype(str).str.split(",").str[0].str.strip() # remove os valores após a virgula
    
    dados = dados_valores.assign(concentracao=500) # insere o valor fixo de 500 em uma coluna

    # Reindexa a tabela, pula 1 linha para cada index para faccilitar a edição visual
    dados.index = range(0, len(dados) * 2, 2)
    dados = dados.reindex(index=range(len(dados) * 2))
    dados = dados.reset_index(drop=True)

    # Insere valores a novas colunas, para fazer cálculos
    col_mgL = "mg/L" if "mg/L" in dados_risco.columns else dados_risco.columns[4]
    col_mgL1 = "mg/L.1" if "mg/L.1" in dados_risco.columns else dados_risco.columns[5]
    dados = dados.assign(efeito=(["C", "NC"] * int(len(dados) / 2)))
    dados = dados.assign(CNCA=dados_risco[col_mgL])
    dados = dados.assign(Aberto="")
    dados = dados.assign(CNCF=dados_risco[col_mgL1])
    dados = dados.assign(Fechado="")

    # Lógica que averigua qual o menor valor entre C, NC para Aberto e Fechado
    for index, row in dados.iterrows():
        # Faz para cada 2 linhas pra pegar as linhas seguintes
        if index % 2 == 0:
            # Podia ser uma função a parte, mas não é
            try: 
                cA = float(dados.loc[index, "CNCA"])
            except:
                cA = 1000000000
            
            try: 
                ncA = float(dados.loc[index+1, "CNCA"])
            except:
                ncA = 1000000000

            # Verifica se os valores existem e se são menores que o VOR, senão substitui
            if cA == 1000000000 and ncA == 1000000000:
                valor_aberto = ""
            else: 
                valor_aberto = min(cA, ncA) # pega o menor
                if valor_aberto < dados.loc[index, "Valor VOR (mg/l)"]:
                    valor_aberto = dados.loc[index, "Valor VOR (mg/l)"]
            
            # Mesma lógica que devia ser uma função a parte, mas não é
            try: 
                cF = float(dados.loc[index, "CNCF"])
            except:
                cF = 1000000000
            
            try: 
                ncF = float(dados.loc[index+1, "CNCF"])
            except:
                ncF = 1000000000

            # Verifica se os valores existem e se são menores que o VOR, senão substitui
            if cF == 1000000000 and ncF == 1000000000:
                valor_fechado = ""
            else: 
                valor_fechado = min(cF, ncF)
                if valor_fechado < dados.loc[index, "Valor VOR (mg/l)"]:
                    valor_fechado = dados.loc[index, "Valor VOR (mg/l)"]

            # Atribui na dataframe
            dados.loc[index, "Aberto"] = valor_aberto
            dados.loc[index, "Fechado"] = valor_fechado

    # Recorta o dataframe para ter apenas as colunas neessárias
    dados = dados[["CAS", "Parâmetro", "efeito", "concentracao", "Valor VOR (mg/l)", "VOR", "CNCA", "Aberto", "CNCF", "Fechado"]]

    return dados

# Conversor de rgb para um valor inteiro
def rgbToInt(rgb):
    """Converte uma cor RGB para um valor inteiro correspondente"""
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    
    return colorInt

# Aplica os estilos nas células do excel
def formatarEstilo(planilha_nova, celulas, linhas_total):
    """Edita o excel para enfeitar as células"""
    # Falha se algo faltar
    if not celulas or not planilha_nova:
        print("Erro ao editar planilhas, cancelando operação.")
        return False
    
    # Para cada célula aplica as propriedades definidas no dicionário
    for celula, propiedades in celulas.items():
        planilha_nova[celula].value = propiedades["titulo"] # texto da célula
        planilha_nova.range(propiedades["mescEcentr"]).merge() # células a mesclar
        planilha_nova[celula].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter # alinha centro horizontal
        planilha_nova[celula].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter # alinha centro vertical

        # Se tiver cor de fondo, aplica
        if "cor" in propiedades:
            planilha_nova.range(propiedades["mescEcentr"]).color = propiedades["cor"]

        # Se tiver o tamanho da coluna
        if "tam" in propiedades:
            planilha_nova[celula].column_width = propiedades["tam"]
        
        # Se tiver borda
        if "borda" in propiedades:
            planilha_nova.range(propiedades["mescEcentr"]).api.Borders(9).Weight = 3
        
    # Tamanhos adicionais
    planilha_nova["D1"].wrap_text = True
    planilha_nova["E1"].column_width = 6
    planilha_nova["F1"].column_width = 24
    planilha_nova["H1"].column_width = 14
    planilha_nova["J1"].column_width = 16

    # Separador cabeçalho/valores
    planilha_nova.range("A7:J7").api.Borders(9).Weight = 4
    planilha_nova.range("A7:J7").api.Borders(9).Color = rgbToInt((245, 201, 174)) #"#F5C9AE"
    
    # Para cada linha da tabela que não for cabeçalho aplica os etilos padrões
    linhas = 8 + linhas_total - 1
    for linha in range(8, linhas + 1, 2):
        # Mescla as células
        planilha_nova.range("A%i:A%i" % (linha, linha+1)).merge()
        planilha_nova.range("B%i:B%i" % (linha, linha+1)).merge()
        planilha_nova.range("D%i:D%i" % (linha, linha+1)).merge()
        planilha_nova.range("E%i:E%i" % (linha, linha+1)).merge()
        planilha_nova.range("F%i:F%i" % (linha, linha+1)).merge()
        planilha_nova.range("H%i:H%i" % (linha, linha+1)).merge()
        planilha_nova.range("J%i:J%i" % (linha, linha+1)).merge()

        # Se o valor VOR for maior de 500 pinta a célula
        if planilha_nova[("H%i" % linha)].value not in ["", None]:
            if float(planilha_nova[("H%i" % linha)].value) > 500:
                planilha_nova.range("H%i:H%i" % (linha, linha+1)).color = "#B1B1B1FF"

        if planilha_nova[("J%i" % linha)].value not in ["", None]:
            if float(planilha_nova[("J%i" % linha)].value) > 500:
                planilha_nova.range("J%i:J%i" % (linha, linha+1)).color = "#B1B1B1FF"

        # Alinha Horizontal
        planilha_nova["A%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["B%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["C%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["C%i" % (linha + 1)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["D%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["E%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["F%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["G%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["G%i" % (linha + 1)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["H%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["I%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["I%i" % (linha + 1)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
        planilha_nova["J%i" % (linha)].api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter

        # Alinha Vertical
        planilha_nova["A%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["B%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["C%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["C%i" % (linha+1)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["D%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["E%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["F%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["G%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["G%i" % (linha+1)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["H%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["I%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["I%i" % (linha+1)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter
        planilha_nova["J%i" % (linha)].api.VerticalAlignment = xw.constants.VAlign.xlVAlignCenter

    planilha_nova.range("E8:E%i" % linhas).font.color = "#E26425"

# Insere o valor do dataframe nas linhas
def inserirValores(dados, planilha_nova):
    """Insere os dados do dataframe na planilha indicada"""
    # Se faltar, falha
    if dados.empty or not planilha_nova:
        print("Erro ao editar planilhas, cancelando operação.")
        return False
    
    # Inserção em uma célula, que já funiona para todas
    planilha_nova["A8"].options(pd.DataFrame, header=None, index=False, expand='table').value = dados 

# Feha corretamente o arquivo
def encerrar(workbook, app, caminho_salvar):
    """Salva e finaliza o workbook e app"""
    workbook.save(caminho_salvar)
    workbook.close()
    app.quit()

    return caminho_salvar

# Chama todas as funções para resolver o problema
def executar(arquivo, arquivo_saida, log=print):
    """Interface com as outras funções"""
    # Estilos das céluulas de cabeçalho
    celulas = {
        "A1": {
            "titulo": "CAS",
            "mescEcentr": "A1:A7",
            "tam": 18
        },
        "B1": {
            "titulo": "Substância Quimica de Interesse",
            "mescEcentr": "B1:B7",
            "tam": 32
        },
        "C1": {
            "titulo": "Efeito",
            "mescEcentr": "C1:C7",
            "tam": 6
        },
        "D1": {
            "titulo": "Concentração de solubilidade",
            "mescEcentr": "D1:D7",
            "tam": 16
        },
        "E1": {
            "titulo": "VOR",
            "mescEcentr": "E1:F7",
        },
        "G1": {
            "titulo": "ÁGUA SUBTERRÂNEA - TRABALHADOR COMERCIAL E INDUSTRIAL",
            "mescEcentr": "G1:J1",
            "tam": 14,
            "borda": True
        },
        "G2": {
            "titulo": "UE - 01A_Raso",
            "mescEcentr": "G2:J2",
            "borda": True
        },
        "G3": {
            "titulo": "Vias de exposição:",
            "cor": "#FEF3CF",
            "mescEcentr": "G3:J3",
            "borda": True
        },
        "G4": {
            "titulo": "Instalação",
            "mescEcentr": "G4:J4",
            "borda": True
        },
        "G5": {
            "titulo": "Ambientes\nAbertos",
            "mescEcentr": "G5:H6",
            "borda": True
        },
        "I5": {
            "titulo": "Ambientes\nFechados",
            "mescEcentr": "I5:J6",
            "tam": 14,
            "borda": True
        },
        "G7": {
            "titulo": "mg/L",
            "mescEcentr": "G7:H7"
        },
        "I7": {
            "titulo": "mg/L",
            "mescEcentr": "I7:J7"
        },
    }

    app = xw.App(visible=False) # criação de uma instância excel

    # Os logs são mostrados ao usuário
    try:
        log("Abrindo workbook...")
        workbook, arquivo = lerWorkbook(app, arquivo)

        log("Lendo planilhas...")
        planilha_valores, planilha_risco, planilhas_existe = lerPlanilhas(workbook)

        log("Criando nova planilha...")
        planilha_nova = criarPlanilha(workbook, planilhas_existe)

        log("Extraindo dados de valore...")
        dados_valores = coletaDados(planilha_valores, arquivo, 0)

        log("Extraindo dados de risco...")
        dados_risco = coletaDados(planilha_risco, arquivo, 6)

        log("Transformando dados...")
        dados = transformarDados(dados_valores, dados_risco)

        log("Completando planilha...")
        inserirValores(dados, planilha_nova)

        log("Formatando estilos...")
        formatarEstilo(planilha_nova, celulas, len(dados))

        log("Salvando em: %s" % os.path.basename(arquivo))
    except Exception as e:
        return False, str(e)

    log("Encerrando...")
    encerrar(workbook, app, arquivo_saida)
    return True, arquivo_saida