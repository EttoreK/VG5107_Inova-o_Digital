import camelot
import pandas as pd
import xlwings as xw

def ajuntarEsquerda(linha):
    ajuntado = linha.dropna().values
    return pd.Series(ajuntado, index=range(len(ajuntado)))

pdf = input("Arraste o arquivo pdf para o terminal: ")

tabelas = camelot.read_pdf(pdf, 
    flavor="stream", 
    pages="1",
)

dados = pd.DataFrame(tabelas[0].df)

linha_corte = dados.loc[lambda df: df[0]=="Parâmetros", :].index[0]
dados_cabecalho = dados.loc[0:linha_corte-1]
dados_analise = dados.loc[linha_corte:]

dados_cabecalho = dados_cabecalho.replace("", pd.NA)
dados_cabecalho.dropna(how="all", axis=1, inplace=True)
dados_cabecalho = dados_cabecalho.apply(ajuntarEsquerda, axis=1)
temp = dados_cabecalho[dados_cabecalho.iloc[0:, 2:].notnull()].dropna(how="all")
temp.dropna(how="all", axis=1, inplace=True)
temp.columns = [0, 1]
dados_cabecalho = dados_cabecalho.iloc[0:, 0:2]
dados_cabecalho = pd.concat([dados_cabecalho, temp])
dados_cabecalho.iloc[0:, 0] = dados_cabecalho.iloc[0:, 0].str.replace(":", "", regex=False).str.strip()
info_cabecalho = dados_cabecalho.set_index(0).to_dict("index")

nome_amostra = info_cabecalho.get("Identificação do Cliente").get(1).replace(" ", "")
data_coleta = info_cabecalho.get("Data da Amostragem").get(1)
data_coleta_dt = pd.to_datetime(data_coleta)
data_coleta_dia = data_coleta_dt.strftime("%m/%d/%Y")
data_coleta_hora = data_coleta_dt.strftime("%H:%M:%S")
id_interna = nome_amostra + "_" + str(int(data_coleta_dt.timestamp()))

titulo = dados_analise.iloc[:2,:]
dados_analise = dados_analise[2:]
colunas_nomes = titulo.agg("sum")
dados_analise.columns = colunas_nomes.str.strip()
dados_analise.index = range(0, len(dados_analise))
dados_analise = dados_analise.reindex(index=range(len(dados_analise)))
dados_analise = dados_analise.reset_index(drop=True)
dados_analise["Diluição"] = dados_analise.iloc[0:, 5]
dados_analise = dados_analise.drop(dados_analise.columns[5], axis=1)
dados_analise.dropna(how="all", axis=1, inplace=True)

dados_analise["Identificação interna"]=id_interna
dados_analise["Nome da amostra"]=nome_amostra
dados_analise["Data de coleta"]=data_coleta_dia
dados_analise["Horário de coleta"]=data_coleta_hora
dados_analise["Unidade"] = dados_analise["Unidade"].str.replace("µ", "u", regex=False)
dados_analise.loc[dados_analise["Resultados analíticos"].str.contains("<"), "Resultados analíticos"] = "< LQ"

dados_final = dados_analise.rename(columns={"Parâmetros": "Parâmetro químico", "Resultados analíticos": "Resultado", "LQ / Faixa": "Limite de Quantificação (LQ)"})
dados_final = dados_final[["Identificação interna", "Nome da amostra", "Data de coleta", "Horário de coleta", "Parâmetro químico", "Resultado", "Unidade", "Limite de Quantificação (LQ)"]]
dados_final.index = range(0, len(dados_final) * 2, 2)
dados_final = dados_final.reindex(index=range(len(dados_final) * 2))
dados_final = dados_final.reset_index(drop=True)
colunas = list(dados_final.columns.values)
# print(dados_cabecalho)
with xw.App(visible=False) as app:
    wb = app.books.add()
    planilha = wb.sheets.add("Relatório")
    planilha["A1"].options(list).value = colunas
    planilha["A4"].options(pd.DataFrame, header=None, index=False, expand="table").value = dados_final

    planilha.range("A1:A3").merge()
    planilha.range("B1:B3").merge()
    planilha.range("C1:C3").merge()
    planilha.range("D1:D3").merge()
    planilha.range("E1:E3").merge()
    planilha.range("F1:F3").merge()
    planilha.range("G1:G3").merge()
    planilha.range("H1:H3").merge()

    for linha in range(4, len(dados_final) + 3, 2):
        planilha.range("A%i:A%i" % (linha, linha+1)).merge()
        planilha.range("B%i:B%i" % (linha, linha+1)).merge()
        planilha.range("C%i:C%i" % (linha, linha+1)).merge()
        planilha.range("D%i:D%i" % (linha, linha+1)).merge()
        planilha.range("E%i:E%i" % (linha, linha+1)).merge()
        planilha.range("F%i:F%i" % (linha, linha+1)).merge()
        planilha.range("G%i:G%i" % (linha, linha+1)).merge()
        planilha.range("H%i:H%i" % (linha, linha+1)).merge()

        if planilha[("F%i" % linha)].value == "< LQ":
            planilha.range("F%i:F%i" % (linha, linha+1)).color = "#d0cece"
    
    # linhas = planilha.range("A4:H4").end("down").row
    linhas = len(dados_final) + 3
    celulas = planilha.range("A1:H%i" % linhas)
    celulas.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    celulas.api.VerticalAlignment = xw.constants.VAlign.xlVAlignTop
    celulas.wrap_text = True
    celulas.column_width = 13
    celulas.api.Borders(11).Weight = 3
    celulas.api.Borders(12).Weight = 3
    # celulas.font.bold = True
    planilha.range("A1:A1").api.VerticalAlignment = xw.constants.VAlign.xlVAlignBottom
    planilha.range("A%i:A%i" % (linhas+1, 1048576)).row_height = 0
    planilha.range("I:XFD").column_width = 0

    wb.save("Relatório.xlsx")
    wb.close()

# print(len(dados_final))