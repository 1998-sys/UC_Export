
import os
import xlwings as xw
from writers.util_writer_oleo import (incerteza_temp_oleo, incerteza_percentual, erro_fiducial, formatar_percentual, 
encontrar_celula_resolucao, encontrar_celula_erro_fiducial, encontrar_celula_incerteza_pressao, encontrar_celula_erro_fiducial_pressao,
encontrar_celula_bsw_maximo, encontrar_celula_incerteza_bsw, encontrar_celula_pressao_op, encontrar_celula_densidade_op, encontrar_celula_temp_op)
from writers.utils_writer import faixas_calibradas, calcular_amplitudes, dados_secundários


def preencher_meter_run_param(wb ,dados):
    incert_transm = incerteza_temp_oleo(dados.get('temperatura'))
    incert_termo = incerteza_temp_oleo(dados.get('termoresistencia'))
    incert_perc_pressao = incerteza_percentual(dados, {"pressao_estatica": None})
    erro_fid = erro_fiducial(dados, {"pressao_estatica": None})
    amplitudes = calcular_amplitudes(faixas_calibradas(dados))
    dados_op = dados.get("dados_fluxo_oleo", {})
    print(f"Incerteza calculada: {incert_transm}")
    print(f"Incerteza calculada: {incert_termo}")
    print(f"Incerteza percentual: {incert_perc_pressao}")
    print(f"Erro fiducial: {erro_fid}")
    print(f"Dados de operação: {dados_op}")

    ws = wb.sheets["Meter run parameters"]

     
    inc_transm = incert_transm.get("maior_incerteza") if incert_transm else None
    inc_termo = incert_termo.get("maior_incerteza") if incert_termo else None
    erro_transm = incert_transm.get("maior_erro") if incert_transm else None
    erro_termo = incert_termo.get("maior_erro") if incert_termo else None

    
    # Escreve incerteza combinada transmissor e termoresistência
    if inc_transm is not None and inc_termo is not None:
        cel_inc = encontrar_celula_resolucao(ws)
        cel_inc.formula = f"=SQRT(({inc_transm}^2)+({inc_termo}^2))"

    # Escreve erro fiducial combinado transmissor e termoresistência
    if erro_transm is not None and erro_termo is not None:
        cel_fid = encontrar_celula_erro_fiducial(ws)
        cel_fid.formula = f"=SQRT(({erro_transm}^2)+({erro_termo}^2))"
    
    # Escreve incerteza percentual pressão estática
    amplitude = amplitudes.get("pressao_estatica")
    inc_perc = incert_perc_pressao.get("pressao_estatica") if incert_perc_pressao else None
    if amplitude is not None:
        cel_incp = encontrar_celula_incerteza_pressao(ws)
        cel_incp.value = f"={amplitude}*{inc_perc}%"

    # Escreve erro fiducial pressão estática
    erro_fidu = erro_fid.get("pressao_estatica") if erro_fid else None
    if erro_fiducial is not None:
        cel_err_p = encontrar_celula_erro_fiducial_pressao(ws)
        cel_err_p.value = f"={amplitude}*{erro_fidu}%"
    
    # Escreve densidade de operação
    densidade_ref = dados_op.get('densidade') if dados_op else None
    if densidade_ref is not None:
        cel_densidade = encontrar_celula_densidade_op(ws)
        cel_densidade.value = densidade_ref
        

    # Escreve temperatura de operação
    temp_ref = dados_op.get('temperatura') if dados_op else None
    if temp_ref is not None:
        cel_tempop=encontrar_celula_temp_op(ws)
        cel_tempop.value = temp_ref   

    # Escreve pressão de operação
    pressao_ref = dados_op.get('pressao') if dados_op else None
    if pressao_ref is not None:
        cel_pop=encontrar_celula_pressao_op(ws)
        cel_pop.value=pressao_ref

    # Escreve BSW máximo permitido
    bsw_max = dados_op.get('bsw_max') if dados_op else None
    bsw_max = formatar_percentual(bsw_max)
    if bsw_max is not None:
        cel_bswm=encontrar_celula_bsw_maximo(ws)
        cel_bswm.value = bsw_max
    
    
    incert_bsw = dados_op.get('incerteza_bsw') if dados_op else None
    incert_bsw = formatar_percentual(incert_bsw)
    if incert_bsw is not None:
        inc_bsw=encontrar_celula_incerteza_bsw(ws)
        inc_bsw.value = incert_bsw


def preencher_equipament_list(wb, dados):
    sec_dados = dados_secundários(dados)
    print(f"Dados secundários: {sec_dados}")

    ws = wb.sheets["Equipment list"]
   
    linhas = {
        "temperatura": 17,
        "termoresistencia": 18,
        "pressao_estatica": 16,
    }

    for instrumento, linha in linhas.items():

        info = sec_dados.get(instrumento, {})

        tag = info.get("tag")
        ns = info.get("numero_serie")
        cert = info.get("certificado")

        
        if tag or ns:

            if tag and ns:
                texto = f"TAG: {tag}\nNS: {ns}"
            elif tag:
                texto = f"TAG: {tag}"
            else:
                texto = f"NS: {ns}"

            celula = ws.range(f"D{linha}")
            celula.value = texto
            celula.api.WrapText = True  

        # 🔹 Certificado na coluna F
        if cert is not None:
            ws.range(f"F{linha}").value = cert


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
        preencher_equipament_list(wb, dados)
    

        app.calculate()

        wb.save()

        wb.close()

    finally:

        app.quit()