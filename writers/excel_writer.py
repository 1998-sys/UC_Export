
import os
import xlwings as xw
from writers.utils_writer import (faixas_calibradas, calcular_amplitudes, formatar_celula_valor, incerteza_absoluta,
                                   erro_fiducial_abs, obter_k, incerteza_temperatura, incert_temp_comb, dados_secundários, dados_placa, obter_respostas,
                                   encontrar_celula)
import re
import shutil

def preencher_gas_parameters(wb, dados):
    amplitudes = calcular_amplitudes(faixas_calibradas(dados))
    incerteza_abs = incerteza_absoluta(dados, amplitudes)
    erro_fid = erro_fiducial_abs(dados, amplitudes)
    k_val = obter_k(dados)
    incert_transm = incerteza_temperatura(dados.get('temperatura'))
    incert_termo = incerteza_temperatura(dados.get('termoresistencia'))
    icert_comb = incert_temp_comb(incert_transm, incert_termo)
    temp_ref = dados.get('dados_operacao', {}).get('temperatura')
    pres_ref = dados.get('dados_operacao', {}).get('pressao')
    

    ws = wb.sheets["Gas parameters"]

    valor = pres_ref
    if valor is not None:
        p_ref = encontrar_celula(ws,"Pressão estática (static pressure), P",coluna_saida="F")
        p_ref.value=valor
        p_ref.api.Locked = True
    
    valor = temp_ref
    if valor is not None:
        t_ref = encontrar_celula(ws,"Temperatura (Temperature), T",coluna_saida="F")
        t_ref.value=valor
        t_ref.api.Locked = True

    valor = amplitudes.get("dpt_alta")
    if valor is not None:
        dpt_alta = encontrar_celula(ws, "Pressão Diferencial Alta (High Differential Pressure)", coluna_saida="F")
        dpt_alta.value = valor
        dpt_alta.api.Locked = True

    valor = amplitudes.get("dp_media")
    if valor is not None:
        dpt_media = encontrar_celula(ws, "Pressão Diferencial Média (Avg Differential Pressure)", coluna_saida="F")
        dpt_media.value = valor
        dpt_media.api.Locked = True

    valor = amplitudes.get("dp_baixa")
    if valor is not None:
        dpt_baixa = encontrar_celula(ws, "Pressão Diferencial Baixa (Low Differential Pressure)", coluna_saida="F")
        dpt_baixa.value = valor
        dpt_baixa.api.Locked = True


    valor = incerteza_abs.get("dpt_alta")
    if valor is not None:
        inc_alta = encontrar_celula(ws, "Pressão diferencial de Alta", coluna_saida="E", tipo_match="exact")
        inc_alta.value = valor
        inc_alta.api.Locked = True
        
    
    valor = erro_fid.get("dpt_alta")
    if valor is not None:
        fid_alta = encontrar_celula(ws, "(High Differential Pressure) Erro Fiducial (Fiducial Error)", coluna_saida="E", tipo_match="exact")
        fid_alta.value = valor
        fid_alta.api.Locked = True

        
    valor = incerteza_abs.get("dp_media")
    if valor is not None:
        inc_media = encontrar_celula(ws, "Pressão diferencial de Média", coluna_saida="E", tipo_match="exact")
        inc_media.value = valor
        inc_media.api.Locked = True
    
    
    valor = erro_fid.get("dp_media")
    if valor is not None:
        fid_medio = encontrar_celula(ws, "(Medium Range Differential Pressure) Fiducial Error", coluna_saida="E", tipo_match="exact")
        fid_medio.value = valor
        fid_medio.api.Locked = True

        
    valor = incerteza_abs.get("dp_baixa")
    if valor is not None:
        inc_baixa = encontrar_celula(ws, "Pressão diferencial de Baixa", coluna_saida="E", tipo_match="exact")
        inc_baixa.value = valor
        inc_baixa.api.Locked = True
        
        
    valor = erro_fid.get("dp_baixa")
    if valor is not None:
        celu_fid_baixa = encontrar_celula(ws, "(Low Range Differential Pressure) Pressure) Fiducial Error", coluna_saida="E", tipo_match="exact")
        celu_fid_baixa.value = valor
        celu_fid_baixa.api.Locked = True
        
    
    
    valor = incerteza_abs.get("pressao_estatica")
    if valor is not None:
        cel_inc_estatica = encontrar_celula(ws, "Pressão estática", coluna_saida="E", tipo_match="exact")
        cel_inc_estatica.value = valor
        cel_inc_estatica.api.Locked = True
        
    

    valor = erro_fid.get("pressao_estatica")
    if valor is not None:
        celu_fid_estatica = encontrar_celula(ws, "(Static Pressure) Erro Fiducial (Fiducial Error)", coluna_saida="E", tipo_match="exact")
        celu_fid_estatica.value = valor
        celu_fid_estatica.api.Locked = True
        

    valor_k = k_val.get("dpt_alta")
    if valor_k is not None:
        cel_k_alta = encontrar_celula(ws, "K factor (Alta)", coluna_busca="G", coluna_saida="G", tipo_match="exact", offset_linha=2)
        cel_k_alta.value=valor_k
        cel_k_alta.api.Locked = True
        
        
    valor_k = k_val.get("dp_media")
    if valor_k is not None:
        cl_k_media = encontrar_celula(ws, "K factor (Média)", coluna_busca="G", coluna_saida="G", tipo_match="exact", offset_linha=2)
        cl_k_media.value = valor_k
        cl_k_media.api.Locked = True

    valor_k = k_val.get("dp_baixa")
    if valor_k is not None:
        cl_k_baixa = encontrar_celula(ws, "K factor (Baixa)", coluna_busca="G", coluna_saida="G", tipo_match="exact", offset_linha=2)
        cl_k_baixa.value = valor_k
        cl_k_baixa.api.Locked = True
        
    
    valor_k = k_val.get("pressao_estatica")
    if valor_k is not None:
        cl_k_estatica = encontrar_celula(ws, "K factor estática", coluna_busca="G", coluna_saida="G", tipo_match="exact", offset_linha=2)
        cl_k_estatica.value = valor_k
        cl_k_estatica.api.Locked = True

    inc_transm = incert_transm.get("incerteza") if incert_transm else None
    k_trasm = incert_transm.get("k") if incert_transm else None
    err_transm = incert_transm.get("erro") if incert_transm else None
    if inc_transm is not None:
        cel_inc_transm = encontrar_celula(ws, "Inc  transm",coluna_busca='X' ,coluna_saida="X", tipo_match="exact", offset_linha=1)
        cel_inc_transm.value = inc_transm
        cel_inc_transm.api.Locked = True
        cel_k_transm = encontrar_celula(ws, "k transm",coluna_busca='Y' ,coluna_saida="Y", tipo_match="exact", offset_linha=1)
        cel_k_transm.value = k_trasm
        cel_k_transm.api.Locked = True     
        cel_err_transm = encontrar_celula(ws, "erro residual",coluna_busca='Z' ,coluna_saida="Z", tipo_match="exact", offset_linha=1)
        cel_err_transm.value = err_transm       
        cel_err_transm.api.Locked = True
        
    
    inc_termo = incert_termo.get("incerteza") if incert_termo else None
    k_termo = incert_termo.get("k") if incert_termo else None
    err_termo = incert_termo.get("erro") if incert_termo else None
    if inc_termo is not None:
        cel_inc_termo = encontrar_celula(ws, "Inc termo",coluna_busca='X' ,coluna_saida="X", tipo_match="exact", offset_linha=1)
        cel_inc_termo.value = inc_termo
        cel_inc_termo.api.Locked = True
        cel_k_termo = encontrar_celula(ws, "k termo",coluna_busca='Y' ,coluna_saida="Y", tipo_match="exact", offset_linha=1)
        cel_k_termo.value = k_termo
        cel_k_termo.api.Locked = True
        cel_err_termo = encontrar_celula(ws, "erro residual",coluna_busca='Z' ,coluna_saida="Z", tipo_match="exact", offset_linha=1)
        cel_err_termo.value = err_termo
        cel_err_termo.api.Locked = True
        
    if icert_comb is not None:
        incert_temp = encontrar_celula(ws, "Temperatura", coluna_saida="E", tipo_match="exact")
        fid_temp = encontrar_celula(ws, "(Temperature) Erro Fiducial (Fiducial Error)", coluna_saida="E", tipo_match="exact")
        k_temp = encontrar_celula(ws, "K factor Temp", coluna_busca="G", coluna_saida="G", tipo_match="exact", offset_linha=2)
        incert_temp.value = icert_comb.get("incerteza")
        incert_temp.api.Locked = True
        fid_temp.value = icert_comb.get("erro")
        fid_temp.api.Locked = True
        k_temp.value = icert_comb.get("k")
        k_temp.api.Locked = True
    
    amplitudes = None
    incerteza_abs = None
    erro_fid = None
    k_val = None
    incert_transm = None
    incert_termo = None
    icert_comb = None
    
def preencher_meter_run_parameter(wb, dados):

    placa_dados = dados_placa(dados)
    ws = wb.sheets["Meter run parameters"]

    diametro = placa_dados.get("diametro_orificio", {}).get("valor", None)
    incert = placa_dados.get("diametro_orificio", {}).get("incerteza", None)
    k_placa = placa_dados.get("diametro_orificio", {}).get('k', None)
    coef_placa = placa_dados.get("coef_dilatacao", None)

    if diametro is not None:
        cel = ws.range("F42")
        cel.value = diametro
    
    if incert is not None:
        cel = ws.range("E49")
        cel.value = incert

    if k_placa is not None:
        cel = ws.range("I49")
        cel.value = k_placa

    if coef_placa is not None:
        cel = ws.range("N44")
        cel.value = coef_placa    
    
    placa_dados= None
    diametro = None
    incert = None
    k_placa = None
    coef_placa = None
     
def preencher_cromatografia(wb, dados):
    print("Entrou na função completa de cromatografia")

    cromatografia = dados.get("cromatografia")

    if not cromatografia:
        print("Nenhum dado de cromatografia encontrado. Mantendo dados anteriores.")
        return

    ws = wb.sheets["Chromatography"]

    ws.range("B2:E200").clear_contents()

    linha = 2

    componentes = [
        c for c in cromatografia.get("componentes", [])
        if (c.get("rotulo") or "").upper() != "H2S"
        and "HIDROG" not in (c.get("nome") or "").upper()
    ]

    
    if not componentes:
        print("Cromatografia existe, mas não possui componentes. Mantendo dados anteriores.")
        return

    
    for comp in componentes:
        rotulo = comp.get("rotulo")
        nome = comp.get("nome")

        ws.range(f"B{linha}").value = rotulo
        ws.range(f"C{linha}").value = nome

        if comp.get("molpct") is not None:
            cel = ws.range(f"D{linha}")
            cel.value = float(comp.get("molpct"))
            formatar_celula_valor(cel)

        if comp.get("incerteza") is not None:
            cel = ws.range(f"E{linha}")
            cel.value = float(comp.get("incerteza"))
            formatar_celula_valor(cel)

        linha += 1

    linha += 1

    
    ws.range(f"B{linha}").value = "Propriedades do Gas - Condição Padrão (1)"
    ws.range(f"B{linha}").api.Font.Bold = True
    ws.range(f"C{linha}").value = "Referência"

    linha += 1

    propriedades_padrao = cromatografia.get("propriedades_condicao_padrao", [])

    for prop in propriedades_padrao:
        ws.range(f"B{linha}").value = prop.get("nome")
        ws.range(f"C{linha}").value = prop.get("referencia")

        if prop.get("valor") is not None:
            cel = ws.range(f"D{linha}")
            cel.value = float(prop.get("valor"))
            formatar_celula_valor(cel)

        if prop.get("incerteza") is not None:
            cel = ws.range(f"E{linha}")
            cel.value = float(prop.get("incerteza"))
            formatar_celula_valor(cel)

        linha += 1

    linha += 1

   
    ws.range(f"B{linha}").value = "Propriedades do Gas - Condições de Amostragem"
    ws.range(f"B{linha}").api.Font.Bold = True
    ws.range(f"C{linha}").value = "Referência"

    linha += 1

    propriedades_amostragem = cromatografia.get(
        "propriedades_condicoes_amostragem", []
    )

    for prop in propriedades_amostragem:
        ws.range(f"B{linha}").value = prop.get("nome")
        ws.range(f"C{linha}").value = prop.get("referencia")

        if prop.get("valor") is not None:
            cel = ws.range(f"D{linha}")
            cel.value = float(prop.get("valor"))
            formatar_celula_valor(cel)

        if prop.get("incerteza") is not None:
            cel = ws.range(f"E{linha}")
            cel.value = float(prop.get("incerteza"))
            formatar_celula_valor(cel)

        linha += 1
    
    cromatografia = None

def preencher_equipament_list(wb, dados):

    sec_dados = dados_secundários(dados)
    placa_dados = dados_placa(dados)
    


    ws = wb.sheets["Equipment List"]

    linhas = {
        "temperatura": 16,
        "termoresistencia": 17,
        "pressao_estatica": 18,
        "dpt_alta": 19,
    }

    for instrumento, linha in linhas.items():

        info = sec_dados.get(instrumento, {})

        tag = info.get("tag")
        ns = info.get("numero_serie")
        cert = info.get("certificado")

        if tag is not None:
            ws.range(f"D{linha}").value = tag
        if ns is not None:
            ws.range(f"E{linha}").value = ns
        if cert is not None:
            ws.range(f"F{linha}").value = cert

    print(placa_dados)
    tag_placa = placa_dados.get("tag")
    ns_placa = placa_dados.get("numero_serie")
    cert_placa = placa_dados.get("certificado")

    if tag_placa is not None:
        ws.range("D15").value = tag_placa

    if ns_placa is not None:
        ws.range("E15").value = ns_placa

    if cert_placa is not None:
        ws.range("F15").value = cert_placa

    sec_dados = None
    placa_dados = None

def preencher_report(wb, dados):

    respostas = obter_respostas()
    ws = wb.sheets["Report"]

    placa = respostas.get("placa", False)
    cromatografia = respostas.get("cromatografia", False)

    secundarios = any([
        respostas.get("dpt_alta"),
        respostas.get("dp_media"),
        respostas.get("dp_baixa"),
        respostas.get("pressao_estatica"),
        respostas.get("temperatura"),
        respostas.get("termoresistencia")
    ])

    if placa and cromatografia and secundarios:
        texto = "Troca de placa de orifício, Atualização de cromatografia e calibração de secundários"

    elif placa and cromatografia:
        texto = "Troca de placa de orifício + Atualização de cromatografia"

    elif placa and secundarios:
        texto = "Troca de placa de orifício + calibração de secundários"

    elif placa:
        texto = "Troca de placa de orifício"

    elif cromatografia:
        texto = "Atualização de cromatografia"

    elif secundarios:
        texto = "Calibração de instrumentos secundários"

    else:
        texto = ""

    ws.range("C13").clear_contents()
    ws.range("C13").value = texto

def _incrementar_nome(caminho_excel):
    """Gera um novo caminho incrementando o último número do nome do arquivo.
    Ex: UCG-FE-3115-03-26-04.xlsx → UCG-FE-3115-03-26-05.xlsx
    """
    pasta = os.path.dirname(caminho_excel)
    nome = os.path.splitext(os.path.basename(caminho_excel))[0]
    ext = os.path.splitext(caminho_excel)[1]

    match = re.search(r'(\d+)(?!.*\d)', nome)
    if match:
        numero = match.group(1)
        novo_numero = str(int(numero) + 1).zfill(len(numero))
        novo_nome = nome[:match.start()] + novo_numero + nome[match.end():]
    else:
        novo_nome = nome + "_1"

    return os.path.join(pasta, novo_nome + ext)


def processar_planilha_gas(caminho_excel, dados):
    """
    Gera uma nova revisão da planilha de CI de gás a partir de um template existente,
    preenchendo todas as abas com os dados extraídos dos certificados XML.

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
    import shutil

    novo_caminho = _incrementar_nome(caminho_excel)
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

        preencher_gas_parameters(wb, dados)
        preencher_meter_run_parameter(wb, dados)
        preencher_cromatografia(wb, dados)
        preencher_equipament_list(wb, dados)
        preencher_report(wb, dados)

        app.calculate()

        wb.save()

        wb.close()

    finally:

        app.quit()