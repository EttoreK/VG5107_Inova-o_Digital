# Módulos necessários
import pandas as pd
import xlwings as xw
import os

# Lê as planilhas do arquivo
def lerPlanilhas(arquivo):
    """Usa uma instância excel para pegar os dados das planilhas"""
    app = xw.App(visible=False) # cria instânia excel
    workbook = app.books.open(arquivo) # abre o arquivo
    planilhas = workbook.sheets # pega as planilhas
    planilhas_arquivo = {p.index: p.name for p in planilhas} # cria um dicionário de planilhas
    workbook.close() # fecha o arquivo
    app.quit() # fecha o app

    return planilhas_arquivo

# Coleta o código da task do relatório
def extrairInfos(arquivo, planilhas_arquivo):
    """Coleta o task_code da planilha"""
    planilha_cadastro = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(1), index_col=False) # lê a planilha 1

    # Limpa os NA do dataframe
    planilha_cadastro.dropna(how="all", axis=1, inplace=True) 
    planilha_cadastro.dropna(how="all", axis=0, inplace=True)

    # Removo caracteres que não fazem sentido em chavevs de um dicionário
    planilha_cadastro.iloc[0:, 0] = planilha_cadastro.iloc[0:, 0].astype(str).str.replace(":", "", regex=False).str.strip()
    planilha_cadastro.iloc[0:, 0] = planilha_cadastro.iloc[0:, 0].str.replace("*", "", regex=False).str.strip()

    # Renomeia as colunas pra ficar melhor no diionário e cria um diionário a partir do dataframe
    planilha_cadastro = planilha_cadastro.rename(columns={planilha_cadastro.columns[0]: "indice", planilha_cadastro.columns[1]: "valor"})
    info_cadastro = planilha_cadastro.set_index("indice").to_dict("index")
    
    # Pega o valor do task_ode ou usa um fallback
    try:
        task_code = info_cadastro.get("Task").get("valor")
    except:
        task_code = "TASK_DESCONHECIDA"

    return task_code

# Manipulação simples para ajeitar a tabela para transformações futuras
def manipularTabela(arquivo, planilhas_arquivo):
    """Normaliza a tabela"""
    # Divide a tabela em 2 partes por causa das células mescladas
    planilha_ezmtp = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(2), index_col=False, skiprows=2)
    planilha_colunas = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(2), index_col=False, skiprows=3, skipfooter=(len(planilha_ezmtp) - 2))
    
    # Renomeia as colunas da tabela que importa com os nomes certos
    colunas_ezmtp = planilha_ezmtp.columns
    if len(colunas_ezmtp) > 12:
        planilha_ezmtp = planilha_ezmtp.rename(columns={
            colunas_ezmtp[8]: planilha_colunas.columns[8],
            colunas_ezmtp[9]: planilha_colunas.columns[9],
            colunas_ezmtp[10]: planilha_colunas.columns[10],
            colunas_ezmtp[11]: planilha_colunas.columns[11],
            colunas_ezmtp[12]: planilha_colunas.columns[12]
        })
    
    planilha_ezmtp.dropna(how="all", axis=1, inplace=True) # limpa os NA a tabela

    # Pega a partir da 2ª coluna
    if planilha_ezmtp.shape[1] > 2:
        planilha_ezmtp = planilha_ezmtp.iloc[1:, 2:]
    
    # Distribui os valores das colunas por cada linha
    planilha_ezmtp = planilha_ezmtp.melt(
        id_vars=["sys_loc_code", "Data da medição", "Hora da medição", "Método de coleta", "Tipo do turbidímetro"],
        var_name="param_code_unit",
        value_name="param_value"
    )

    planilha_ezmtp[["param_code", "param_unit"]] = planilha_ezmtp["param_code_unit"].str.split("(", n=1, expand=True) # separa o parametro da unidade

    # Cria um datetime completo
    data_completa = planilha_ezmtp["Data da medição"].astype(str) + " " + planilha_ezmtp["Hora da medição"].astype(str)
    planilha_ezmtp["measurement_date"] = pd.to_datetime(data_completa, errors='coerce').dt.strftime("%d/%m/%Y %H:%M")

    # Renomeia mais colunas
    planilha_ezmtp = planilha_ezmtp.rename(columns={
        planilha_ezmtp.columns[0]: "#sys_loc_code",
        planilha_ezmtp.columns[3]: "measurement_method",
        planilha_ezmtp.columns[4]: "remark",
    })

    return planilha_ezmtp

# Reformula a tabela para deixar da forma desejada
def formatarTabela(planilha_ezmtp, task_code):
    """Faz a tabela ficar como especificado na imagem"""
    # Pega as olunas neessárias e cria a coluna om o task_code
    colunas_finais = ["#sys_loc_code", "param_code", "param_value", "param_unit", "measurement_method", "measurement_date", "remark"]
    planilha_ezmtp = planilha_ezmtp[[c for c in colunas_finais if c in planilha_ezmtp.columns]]
    planilha_ezmtp = planilha_ezmtp.assign(task_code=task_code)

    # Substitui os nomes dos parâmetros por seus códigos
    repor_cods = {"Temp.": "Temp", "Condutividade": "Cond elet", "Turbidez": "Turb", "Responsável pela coleta": "Resp Colet", "Condição climática": "Cond clim"}
    planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].astype(str).str.rstrip(" \n")
    planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace(repor_cods)#.str.replace("\n", "").str.strip()

    # Substitui as unidades dos parâmetros por suas formas corretas
    repor_unidad = {"°C)": "C", "µS/cm)": "uS/cm", "NTU)": "ntu", "mg/L)": "mg/l", "mV)": "mv", "None": "-"}
    if "param_unit" in planilha_ezmtp.columns:
        planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].astype(str).replace(repor_unidad).str.strip()

    planilha_ezmtp = planilha_ezmtp.fillna(value="-") # Substitui NA por -

    # Ordena o dataframe por [A-Z] > [a-z] > [0-9]
    planilha_ezmtp = planilha_ezmtp.sort_values(
        by=["#sys_loc_code","param_value"], 
        key=lambda col: col.str.lower()
    )

    return planilha_ezmtp

# Aplica o requerido para as células do excel
def enfeitarSaida(planilha_ezmtp, caminho_salvar):
    """Edita o excel para enfeitar as células"""
    with xw.App(visible=False) as app:
        wb = app.books.add()
        nome_planilha = "Relatório"
        planilha = wb.sheets.add(nome_planilha)
        
        celulas = planilha.range("A1:H1") # células do cabeçalho
        planilha["A1"].options(list).value = list(planilha_ezmtp.columns.values) # textos do cabeçalho
        planilha["A2"].options(pd.DataFrame, header=False, index=False, expand="table").value = planilha_ezmtp # valores do dataframe
        
        # Esconde todas as demais células
        planilha.range("A%i:A1048576" % (len(planilha_ezmtp)+2)).row_height = 0
        planilha.range("I:XFD").column_width = 0

        # Estilos
        celulas.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter # centraliza horizontalmente
        celulas.api.Borders(11).Weight = 3 # tipo de borda e grossura
        planilha.range("A1:H2").api.Borders(12).Weight = 3 # tipo de borda e grossura
        celulas.column_width = 18 # tamanho da célula

        # Verifica se já existe
        if os.path.exists(caminho_salvar):
            # Remove se já existir para evitar erro
            try: 
                os.remove(caminho_salvar)
            except: 
                pass
        
        # Excluí tudo que não for a planilha relatório
        for planilha in wb.sheetsif:
            if planilha.name != nome_planilha: planilha.delete()
        
        wb.save(caminho_salvar) # salva
        wb.close() # fecha

    return caminho_salvar
    
# Chama todas as funções para resolver o problema
def executar(arquivo, arquivo_saida, log=print):
    """Interface com as outras funções"""
    # Os logs são mostrados ao usuário
    try:
        log("Lendo planilhas...")
        planilhas = lerPlanilhas(arquivo)
        
        log("Extraindo informações...")
        task_code = extrairInfos(arquivo, planilhas)
        
        log("Código Task: %s" % task_code)
        log("Manipulando tabela...")
        tabela = manipularTabela(arquivo, planilhas)
        
        log("Formatando dados...")
        tabela_fmt = formatarTabela(tabela, task_code)
        
        log("Gerando Excel final...")
        caminho_final = enfeitarSaida(tabela_fmt, arquivo_saida)
        
        log("Salvando em: %s" % os.path.basename(caminho_final))
        return True, caminho_final
    except Exception as e:
        return False, str(e)