
import os
import xlwings as xw
from writers.util_writer_oleo import incerteza_temp_oleo, incerteza_percentual, erro_fiducial
from writers.utils_writer import faixas_calibradas, calcular_amplitudes



def preencher_meter_run_param(wb ,dados):
    incert_transm = incerteza_temp_oleo(dados.get('temperatura'))
    incert_termo = incerteza_temp_oleo(dados.get('termoresistencia'))
    incert_perc_pressao = incerteza_percentual(dados, {"pressao_estatica": None})
    erro_fid = erro_fiducial(dados, {"pressao_estatica": None})
    amplitudes = calcular_amplitudes(faixas_calibradas(dados))
    print(f"Incerteza calculada: {incert_transm}")
    print(f"Incerteza calculada: {incert_termo}")
    print(f"Incerteza percentual: {incert_perc_pressao}")
    print(f"Erro fiducial: {erro_fid}")

    ws = wb.sheets["Meter run parameters"]

     
    inc_transm = incert_transm.get("maior_incerteza") if incert_transm else None
    inc_termo = incert_termo.get("maior_incerteza") if incert_termo else None
    erro_transm = incert_transm.get("maior_erro") if incert_transm else None
    erro_termo = incert_termo.get("maior_erro") if incert_termo else None

    
    if inc_transm is not None and inc_termo is not None:
        cel = ws.range('M107')
        cel.formula = f"=SQRT(({inc_transm}^2)+({inc_termo}^2))"

    
    if erro_transm is not None and erro_termo is not None:
        cel = ws.range('E117')
        cel.formula = f"=SQRT(({erro_transm}^2)+({erro_termo}^2))"
    

    amplitude = amplitudes.get("pressao_estatica")
    inc_perc = incert_perc_pressao.get("pressao_estatica") if incert_perc_pressao else None
    if amplitude is not None:
        ws.range("E159").value =  f"={amplitude}*{inc_perc}%"

    erro_fidu = erro_fid.get("pressao_estatica") if erro_fid else None
    if erro_fiducial is not None:
        ws.range("E170").value = f"={amplitude}*{erro_fidu}%"
        
       


def processar_planilha_oleo(caminho_excel, dados):

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:

        wb = app.books.open(
            caminho_excel,
            update_links=False,
            read_only=False,
            ignore_read_only_recommended=True
        )

        preencher_meter_run_param(wb, dados)
    

        app.calculate()

        wb.save()

        wb.close()

    finally:

        app.quit()