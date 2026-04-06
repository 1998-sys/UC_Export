
import os
import xlwings as xw
from writers.util_writer_oleo import incerteza_temp_oleo, incerteza_percentual, erro_fiducial, formatar_percentual
from writers.utils_writer import faixas_calibradas, calcular_amplitudes, dados_secundários, encontrar_celula, incrementar_nome
import shutil


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
        cel_inc = encontrar_celula(ws, "Resolução da Termoresistência (Termoresistance resolution)", coluna_saida="M")
        print(f"celula da incerteza combinada: {cel_inc.address}")
        cel_inc.formula_local = f"=RAIZ(SOMAQUAD({inc_transm};{inc_termo}))"

    # Escreve erro fiducial combinado transmissor e termoresistência
    if erro_transm is not None and erro_termo is not None:
        cel_fid = encontrar_celula(ws, "C2.2.2 - Erro Fiducial (Fiducial Error)", coluna_saida="E")
        print(f"celula do erro fiducial combinado: {cel_fid.address}")
        cel_fid.formula_local = f"=RAIZ(SOMAQUAD({erro_transm};{erro_termo}))"

    # Escreve incerteza percentual pressão estática
    amplitude = amplitudes.get("pressao_estatica")
    inc_perc = incert_perc_pressao.get("pressao_estatica") if incert_perc_pressao else None
    if amplitude is not None:
        cel_incp = encontrar_celula(ws, "Incerteza da calibração do medidor de pressão (Pressure meter calibration uncertainty)", coluna_saida="E")
        print(f"celula da incerteza percentual: {cel_incp.address}")
        cel_incp.value = f"={amplitude}*{inc_perc}%"

    # Escreve erro fiducial pressão estática
    erro_fidu = erro_fid.get("pressao_estatica") if erro_fid else None
    if erro_fidu is not None:
        cel_err_p = encontrar_celula(ws, "C3.1.2 - Erro Fiducial (Fiducial Error)", coluna_saida="E")
        print(f"celula do erro fiducial da pressão estática: {cel_err_p.address}")
        cel_err_p.value = f"={amplitude}*{erro_fidu}%"

    # Escreve densidade de operação
    densidade_ref = dados_op.get('densidade') if dados_op else None
    if densidade_ref is not None:
        cel_densidade = encontrar_celula(ws, "Densidade nas condições De Referência (Standard Density), ρ", coluna_saida="F")
        print(f"celula da densidade de operação: {cel_densidade.address}")
        cel_densidade.value = densidade_ref

    # Escreve temperatura de operação
    temp_ref = dados_op.get('temperatura') if dados_op else None
    if temp_ref is not None:
        cel_tempop = encontrar_celula(ws, "Temp. da Termoresistência (Termoresistance temp.) - Ta", coluna_saida="F")
        print(f"celula da temperatura de operação: {cel_tempop.address}")
        cel_tempop.value = temp_ref

    # Escreve pressão de operação
    pressao_ref = dados_op.get('pressao') if dados_op else None
    if pressao_ref is not None:
        cel_pop = encontrar_celula(ws, "Pressão estática (static pressure), P", coluna_saida="F")
        print(f"celula da pressão de operação: {cel_pop.address}")
        cel_pop.value = pressao_ref

    # Escreve BSW máximo permitido
    bsw_max = dados_op.get('bsw_max') if dados_op else None
    bsw_max = formatar_percentual(bsw_max)
    if bsw_max is not None:
        cel_bswm = encontrar_celula(ws, "BSW Máximo  (Max BSW Allowed)", coluna_saida="F")
        print(f"celula do BSW máximo permitido: {cel_bswm.address}")
        cel_bswm.value = bsw_max

    # Escreve incerteza BSW
    incert_bsw = dados_op.get('incerteza_bsw') if dados_op else None
    incert_bsw = formatar_percentual(incert_bsw)
    if incert_bsw is not None:
        cel_inc_bsw = encontrar_celula(ws, "C5.1 Incerteza padrão combinada - BSW (BSW Combined Uncertainty)", coluna_saida="E")
        print(f"celula da incerteza BSW: {cel_inc_bsw.address}")
        cel_inc_bsw.value = incert_bsw


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
    """
    Gera uma nova revisão da planilha de CI de óleo a partir de um template existente,
    preenchendo as abas com os dados extraídos dos certificados XML.

    O arquivo de origem nunca é alterado. Uma cópia com nome incrementado é criada
    antes de qualquer escrita (ex: *-04.xlsx → *-05.xlsx), garantindo rastreabilidade
    de revisões e integridade do template.

    Args:
        caminho_excel (str): Caminho absoluto da planilha de referência (revisão anterior).
        dados (dict): Dados consolidados dos instrumentos, incluindo XMLs parseados,
                      condições operacionais e resultados calculados.

    Raises:
        Exception: Propaga qualquer exceção do xlwings; a instância do Excel é
                   encerrada via `finally` independentemente do resultado.
    """
    novo_caminho = incrementar_nome(caminho_excel)
    shutil.copy2(caminho_excel, novo_caminho)

    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False

    try:

        wb = app.books.open(
            novo_caminho,
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