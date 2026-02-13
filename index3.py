import pandas as pd
import xlwings as xw
import os as os

def eNum(num):
    try: 
        float(num)
        return True
    except:
        return False

app = xw.App(visible=False)

arquivo = input("Arraste o arquivo para dentro do terminal: ")
print("Processando arquivo: %s" % arquivo)

arquivo = arquivo.replace("\\ ", " ").strip()
if not os.path.exists(arquivo):
    print("Falha ao ler arquivo, cancelando operação.")

try:
    os.rename(arquivo, arquivo)
except:
    print("O arquivo está aberto em outro processo, cancelando operação.")

workbook = app.books.open(arquivo)

planilhas = workbook.sheets

planilhas_arquivo = {}
for planilha in planilhas:
    indice = planilha.index
    nome = planilha.name
    planilhas_arquivo[indice] = nome

workbook.close()
app.quit()

planilha_cadastro = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(1), index_col=False)
planilha_cadastro.dropna(how="all", axis=1, inplace=True)
planilha_cadastro.dropna(how="all", axis=0, inplace=True)
planilha_cadastro.iloc[0:, 0] = planilha_cadastro.iloc[0:, 0].str.replace(":", "", regex=False).str.strip()
planilha_cadastro.iloc[0:, 0] = planilha_cadastro.iloc[0:, 0].str.replace("*", "", regex=False).str.strip()
planilha_cadastro = planilha_cadastro.rename(columns={planilha_cadastro.columns[0]: "indice", planilha_cadastro.columns[1]: "valor"})
info_cadastro = planilha_cadastro.set_index("indice").to_dict("index")
task_code = info_cadastro.get("Task").get("valor")

planilha_ezmtp = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(2), index_col=False, skiprows=2)
planilha_colunas = pd.read_excel(arquivo, sheet_name=planilhas_arquivo.get(2), index_col=False, skiprows=3, skipfooter=(len(planilha_ezmtp) - 2))
planilha_ezmtp = planilha_ezmtp.rename(columns={
    planilha_ezmtp.columns[8]: planilha_colunas.columns[8],
    planilha_ezmtp.columns[9]: planilha_colunas.columns[9],
    planilha_ezmtp.columns[10]: planilha_colunas.columns[10],
    planilha_ezmtp.columns[11]: planilha_colunas.columns[11],
    planilha_ezmtp.columns[12]: planilha_colunas.columns[12]
})
planilha_ezmtp.dropna(how="all", axis=1, inplace=True)
planilha_ezmtp = planilha_ezmtp.iloc[1:, 2:]
planilha_ezmtp = planilha_ezmtp.melt(
    id_vars=["sys_loc_code", "Data da medição", "Hora da medição", "Método de coleta", "Tipo do turbidímetro"],
    var_name="param_code_unit",
    value_name="param_value"
)

planilha_ezmtp[["param_code", "param_unit"]] = planilha_ezmtp["param_code_unit"].str.split(" \n", n=1, expand=True)
planilha_ezmtp["measurement_date"] = (
    (planilha_ezmtp["Data da medição"].astype(str)) + " " +
    (planilha_ezmtp["Hora da medição"].astype(str))
).apply(pd.to_datetime)

planilha_ezmtp["measurement_date"] = planilha_ezmtp["measurement_date"].dt.strftime("%d/%m/%Y %H:%M")

planilha_ezmtp = planilha_ezmtp.rename(columns={
    planilha_ezmtp.columns[0]: "#sys_loc_code",
    planilha_ezmtp.columns[3]: "measurement_method",
    planilha_ezmtp.columns[4]: "remark",
})
planilha_ezmtp = planilha_ezmtp[["#sys_loc_code", "param_code", "param_value", "param_unit", "measurement_method", "measurement_date", "remark"]]
planilha_ezmtp = planilha_ezmtp.assign(task_code=task_code)

planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace("Temp.", "Temp")
planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace("Condutividade", "Cond elet")
planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace("Turbidez", "Turb")
planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace("Responsável pela coleta", "Resp Colet")
planilha_ezmtp["param_code"] = planilha_ezmtp["param_code"].replace("Condição climática", "Cond clim")
planilha_ezmtp["param_code"].str.replace("\n", "")
planilha_ezmtp["param_code"].str.strip()

planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].replace("(°C)", "C")
planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].replace("(µS/cm)", "uS/cm")
planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].replace("(NTU)", "ntu")
planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].replace("(mg/L)", "mg/l")
planilha_ezmtp["param_unit"] = planilha_ezmtp["param_unit"].replace("(mV)", "mv")
planilha_ezmtp["param_unit"].str.strip()

planilha_ezmtp = planilha_ezmtp.fillna(value="-")

planilha_ezmtp = planilha_ezmtp.sort_values(
    by=["#sys_loc_code","param_value"], 
    key=lambda col: col.str.lower()
)

with xw.App(visible=False) as app:
    wb = app.books.add()
    planilha = wb.sheets.add("Relatório")
    celulas = planilha.range("A1:H1")
    planilha["A1"].options(list).value = list(planilha_ezmtp.columns.values)
    planilha["A2"].options(pd.DataFrame, header=None, index=False, expand="table").value = planilha_ezmtp
    planilha.range("A%i:A%i" % (len(planilha_ezmtp)+1, 1048576)).row_height = 0
    planilha.range("I:XFD").column_width = 0
    celulas.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    celulas.api.Borders(11).Weight = 3
    planilha.range("A1:H2").api.Borders(12).Weight = 3
    celulas.column_width = 18

    wb.save("Relatório.xlsx")
    wb.close()

# print(planilha_ezmtp)

