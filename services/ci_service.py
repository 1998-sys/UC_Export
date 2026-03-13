from writers.excel_writer import processar_planilha
from openpyxl import load_workbook

def executar_fluxo(ci_path: str, dados: dict):
    processar_planilha(
        caminho_excel=ci_path,
        dados=dados
    )



def identificar_tipo_ci(caminho_excel):

    wb = load_workbook(caminho_excel, read_only=True)

    abas = set(wb.sheetnames)

    wb.close()

    abas_gas = {
        "Report",
        "Gas parameters",
        "Meter run parameters",
        "Coef Disc & Expansao"
    }

    if abas_gas.issubset(abas):
        return "gas"

    return "nao_suportado"