# Módulos necessários
import camelot
import pandas as pd
import xlwings as xw
import os

# Empurra todos os valores da tabela para as colunas mais esquerda, em suas respectivas linhas
def ajuntarEsquerda(linha):
    """Transfere os valores para as colunas mais a esquerda"""
    ajuntado = linha.dropna().values
    return pd.Series(ajuntado, index=range(len(ajuntado)))

# Lê as tabelas do PDF
def lerArquivo(arquivo):
    """Extrai a tabela da página 1"""
    tabelas = camelot.read_pdf(arquivo, 
        flavor="stream", # detecta automaticamente os divisores das tabelas
        pages="1", # Só aceita tabelas da página 1
    )

    return tabelas

# Separa as tabelas de cabeçalho e valores
def extrairTabelas(tabelas):
    """Separa as tabelas e transforma em dataframes"""
    dados = pd.DataFrame(tabelas[0].df) # tabelas[0] já é todo o conteúdo

    # Localiza o termo "Parâmetros" para achar o real cabeçalho da tabela com valores
    linha_corte = dados.loc[lambda df: df[0]=="Parâmetros", :].index[0]
    dados_cabecalho = dados.loc[0:linha_corte-1]
    dados_analise = dados.loc[linha_corte:]

    return dados_cabecalho, dados_analise

# Coleta as informações do cabeçalho e transforma em dicionário
def extrairInfos(dados_cabecalho):
    """Coleta informações relevantes do cabeçalho da página"""
    # Limpa e formata a tabela cabeçalho
    dados_cabecalho = dados_cabecalho.replace("", pd.NA)
    dados_cabecalho.dropna(how="all", axis=1, inplace=True)
    dados_cabecalho = dados_cabecalho.apply(ajuntarEsquerda, axis=1)

    # Tabela temporária para remover os únicos valores que compõem as colunas 3 e 4
    temp = dados_cabecalho[dados_cabecalho.iloc[0:, 2:].notnull()].dropna(how="all")
    temp.dropna(how="all", axis=1, inplace=True)
    temp.columns = [0, 1]

    # Atribuição dos valores à tabela original e redimensionamento da tabela
    dados_cabecalho = dados_cabecalho.iloc[0:, 0:2]
    dados_cabecalho = pd.concat([dados_cabecalho, temp])
    dados_cabecalho.iloc[0:, 0] = dados_cabecalho.iloc[0:, 0].str.replace(":", "", regex=False).str.strip()

    # Dicionário criado da tabela para pegar os valore facilmente
    info_cabecalho = dados_cabecalho.set_index(0).to_dict("index")

    # Coleta de valores relevantes
    nome_amostra = info_cabecalho.get("Identificação do Cliente").get(1).replace(" ", "")
    data_coleta = info_cabecalho.get("Data da Amostragem").get(1)

    # Controle se não der pra converter a data em tempo
    try:
        data_coleta_dt = pd.to_datetime(data_coleta)
        data_coleta_dia = data_coleta_dt.strftime("%m/%d/%Y")
        data_coleta_hora = data_coleta_dt.strftime("%H:%M:%S")
        # id_interna = nome_amostra + "_" + str(int(data_coleta_dt.timestamp()))
        # Converte a data_oleta para um padrão de 5 números que representam os dias no excel
        dias_excel = (data_coleta_dt - pd.to_datetime("1899-12-30")).days
        id_interna = nome_amostra + "_" + str(dias_excel)
    except:
        data_coleta_dt = ""
        data_coleta_dia = ""
        data_coleta_hora = ""
        id_interna = nome_amostra + "_0"
    
    # Retorna um dicionário mais fácil e só com as informações úteis
    return {"nome": nome_amostra, "data": data_coleta, "data_dia": data_coleta_dia, "data_hora": data_coleta_hora, "idi": id_interna}

# Manipulação simples para ajeitar a tabela para transformações futuras
def manipularTabela(dados_analise):
    """Normaliza a tabela"""
    titulo = dados_analise.iloc[:2,:] # pega a primeira linha da tabela
    dados_analise = dados_analise[2:] # remove a primeira linha da tabela
    colunas_nomes = titulo.agg("sum") # concatena todas as strings, para eliminar espaços
    dados_analise.columns = colunas_nomes.str.strip() # separa os tiitulos, antes concatenados, agora normalizados

    # Reindexa a tabela
    dados_analise.index = range(0, len(dados_analise)) 
    dados_analise = dados_analise.reindex(index=range(len(dados_analise)))
    dados_analise = dados_analise.reset_index(drop=True)

    # Coloca os valores da coluna sem titulo na coluna sem valores e destroi o que ficou vazio
    dados_analise["Diluição"] = dados_analise.iloc[0:, 5]
    dados_analise = dados_analise.drop(dados_analise.columns[5], axis=1)
    dados_analise.dropna(how="all", axis=1, inplace=True)

    return dados_analise

# Reformula a tabela para deixar da forma desejada
def formatarTabela(dados_analise, dict_cabecalho):
    """Faz a tabela ficar como especificado na imagem"""
    # Cria colunas com os valores do cabeçalho
    dados_analise["Identificação interna"]=dict_cabecalho.get("idi")
    dados_analise["Nome da amostra"]=dict_cabecalho.get("nome")
    dados_analise["Data de coleta"]=dict_cabecalho.get("data_dia")
    dados_analise["Horário de coleta"]=dict_cabecalho.get("data_hora")

    # Repõe valores baseados em regras de negócio
    dados_analise["Unidade"] = dados_analise["Unidade"].str.replace("µ", "u", regex=False)
    dados_analise.loc[dados_analise["Resultados analíticos"].str.contains("<"), "Resultados analíticos"] = "< LQ"

    # Renomeia o dataframe e recorta o que não importa para o relatório
    dados_final = dados_analise.rename(columns={"Parâmetros": "Parâmetro químico", "Resultados analíticos": "Resultado", "LQ / Faixa": "Limite de Quantificação (LQ)"})
    dados_final = dados_final[["Identificação interna", "Nome da amostra", "Data de coleta", "Horário de coleta", "Parâmetro químico", "Resultado", "Unidade", "Limite de Quantificação (LQ)"]]

    # Reindexa a tabela, pula 1 linha para cada index para faccilitar a edição visual
    dados_final.index = range(0, len(dados_final) * 2, 2)
    dados_final = dados_final.reindex(index=range(len(dados_final) * 2))
    dados_final = dados_final.reset_index(drop=True)

    colunas = list(dados_final.columns.values) # pega os nomes das colunas para facilitar a edição visual

    return dados_final, colunas

# Aplica o requerido para as células do excel
def enfeitarSaida(dados_final, colunas, caminho_salvar):
    """Edita o excel para enfeitar as células"""
    # Abre o app excel para alterar facilmente
    with xw.App(visible=False) as app:
        wb = app.books.add() # cria um novo arquivo
        nome_planilha = "Relatório"
        planilha = wb.sheets.add(nome_planilha) # cria uma nova planilha
        planilha["A1"].options(list).value = colunas # cria todo o cabeçalho

        # Insere todos os dados do dataframe 
        planilha["A4"].options(pd.DataFrame, header=None, index=False, expand="table").value = dados_final

        # Faz a mesclágem das células de cabeçalho
        planilha.range("A1:A3").merge()
        planilha.range("B1:B3").merge()
        planilha.range("C1:C3").merge()
        planilha.range("D1:D3").merge()
        planilha.range("E1:E3").merge()
        planilha.range("F1:F3").merge()
        planilha.range("G1:G3").merge()
        planilha.range("H1:H3").merge()

        # faz a mesclágem das células da tabela
        for linha in range(4, len(dados_final) + 3, 2):
            planilha.range("A%i:A%i" % (linha, linha+1)).merge()
            planilha.range("B%i:B%i" % (linha, linha+1)).merge()
            planilha.range("C%i:C%i" % (linha, linha+1)).merge()
            planilha.range("D%i:D%i" % (linha, linha+1)).merge()
            planilha.range("E%i:E%i" % (linha, linha+1)).merge()
            planilha.range("F%i:F%i" % (linha, linha+1)).merge()
            planilha.range("G%i:G%i" % (linha, linha+1)).merge()
            planilha.range("H%i:H%i" % (linha, linha+1)).merge()
            
            # Pinta a célula se tiver o valor
            if planilha[("F%i" % linha)].value == "< LQ":
                planilha.range("F%i:F%i" % (linha, linha+1)).color = "#d0cece"

        # Pega todas as células da tabela para fazer uma formatação generalizada
        linhas = len(dados_final) + 3
        celulas = planilha.range("A1:H%i" % linhas)

        # Formatação generalizada
        celulas.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter # alinha no centro horizontal
        celulas.api.VerticalAlignment = xw.constants.VAlign.xlVAlignTop # põe o texto no topo
        celulas.wrap_text = True # permite o texto pular linha dentro da célula
        celulas.column_width = 13 # tamanho da célula
        celulas.api.Borders(11).Weight = 3 # tipo de bora e grossura
        celulas.api.Borders(12).Weight = 3
        
        # Formatação do cabeçalho
        planilha.range("A1:A1").api.VerticalAlignment = xw.constants.VAlign.xlVAlignBottom

        # Esconde todas as outras células
        planilha.range("A%i:A%i" % (linhas+1, 1048576)).row_height = 0 
        planilha.range("I:XFD").column_width = 0

        # Salva por cima
        if os.path.exists(caminho_salvar):
            # Remove se já existir para evitar erro
            try: 
                os.remove(caminho_salvar)
            except: 
                pass
        
        # Excluí tudo que  não for a planilha relatório
        for planilha in wb.sheets:
            if planilha.name != nome_planilha: planilha.delete()
        
        # Salva e fecha
        wb.save(caminho_salvar)
        wb.close()

        return caminho_salvar

# Chama todas as funções para resolver o problema
def executar(arquivo, arquivo_saida, log=print):
    """Interface com as outras funções"""
    # Os logs são mostrados ao usuário
    try:
        log("Lendo planilhas...")
        tabelas = lerArquivo(arquivo)

        log("Extraindo tabelas...")
        cabecalho, analise = extrairTabelas(tabelas)
        
        log("Extraindo informações...")
        dict_cabecalho = extrairInfos(cabecalho)
        
        log("Manipulando tabela...")
        dados = manipularTabela(analise)
        
        log("Formatando dados...")
        tabela_final, colunas = formatarTabela(dados, dict_cabecalho)
        
        log("Gerando Excel final...")
        caminho_final = enfeitarSaida(tabela_final, colunas, arquivo_saida)
        
        log("Salvando em: %s" % os.path.basename(caminho_final))
        return True, caminho_final
    except Exception as e:
        return False, str(e)