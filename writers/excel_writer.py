
import os
import xlwings as xw
from writers.utils_writer import (faixas_calibradas, calcular_amplitudes, formatar_celula_valor, incerteza_absoluta,
                                   erro_fiducial_abs, obter_k, incerteza_temperatura, incert_temp_comb, dados_secundários, dados_placa, obter_respostas,
                                   encontrar_celula, incrementar_nome, alterar_ncalculo)
import re
import shutil



def preencher_gas_parameters(wb, dados):
    """
    Preenche a aba "Gas parameters" com os resultados metrológicos calculados
    a partir dos certificados de calibração dos instrumentos de pressão e temperatura.

    Valores escritos (todos bloqueados contra edição após escrita):
      - Condições operacionais de referência: pressão estática e temperatura.
      - Amplitudes das faixas calibradas: DPT Alta, Média e Baixa.
      - Incerteza absoluta e erro fiducial absoluto para cada DPT e pressão estática.
      - Fator K de cobertura para cada instrumento de pressão.
      - Incerteza, K e erro residual individuais do transmissor de temperatura e da
        termoresistência (células auxiliares para composição RSS).
      - Incerteza combinada, erro fiducial combinado e K da temperatura (RSS).

    Args:
        wb: Workbook xlwings aberto da planilha de CI.
        dados (dict): Dados consolidados contendo instrumentos calibrados,
                      faixas, pontos de calibração e condições operacionais.
    """
    amplitudes = calcular_amplitudes(faixas_calibradas(dados))
    incerteza_abs = incerteza_absoluta(dados, amplitudes)
    erro_fid = erro_fiducial_abs(dados, amplitudes)
    k_val = obter_k(dados)
    incert_transm = incerteza_temperatura(dados.get('temperatura'))
    incert_termo = incerteza_temperatura(dados.get('termoresistencia'))
    icert_comb = incert_temp_comb(incert_transm, incert_termo)
    temp_ref = dados.get('dados_operacao', {}).get('temperatura')
    pres_ref = dados.get('dados_operacao', {}).get('pressao')
    cromatografia = dados.get("cromatografia")
    props_padrao  = (cromatografia or {}).get("propriedades_condicao_padrao", {})
    props_amostr  = (cromatografia or {}).get("propriedades_condicoes_amostragem", {})

    prop_z   = props_amostr.get("Fator de Compressibilidade", {})
    prop_mw  = props_padrao.get("Peso Molecular Total (g/mol)", {})
    prop_rho = props_amostr.get("Densidade (kg/m³)", {})

    ws = wb.sheets["Parameters"]

    valor = pres_ref
    if valor is not None:
        p_ref = encontrar_celula(ws,"Pressão estática (static pressure), P",coluna_saida="D")
        p_ref.value=valor
        p_ref.api.Locked = True
    
    valor = temp_ref
    if valor is not None:
        t_ref = encontrar_celula(ws,"Temperatura (Temperature), T",coluna_saida="D")
        t_ref.value=valor
        t_ref.api.Locked = True

    valor = prop_z.get("valor")
    if valor is not None:
        z_cel = encontrar_celula(ws, "Fator de compressibilidade, Z", coluna_saida="D", tipo_match="exact")
        z_cel.value = valor
        z_cel.api.Locked = True
    
    valor = prop_z.get("incerteza")
    if valor is not None:
        z_inc_cel = encontrar_celula(ws, "Fator de compressibilidade, Z (Incert)", coluna_busca='F', coluna_saida="H", tipo_match="exact")
        z_inc_cel.value = valor
        z_inc_cel.api.Locked = True
        
    valor = prop_mw.get("valor")
    if valor is not None:
        mw_cel = encontrar_celula(ws, "Massa molar, M", coluna_saida="D", tipo_match="exact")
        mw_cel.value = valor
        mw_cel.api.Locked = True
    
    valor = prop_mw.get("incerteza")
    if valor is not None:
        mw_inc_cel = encontrar_celula(ws, "Massa molar, M (Incert)", coluna_busca='F', coluna_saida="H", tipo_match="exact")
        mw_inc_cel.value = valor
        mw_inc_cel.api.Locked = True
    
    valor = prop_rho.get("valor")
    if valor is not None:
        rho_cel = encontrar_celula(ws, "Densidade absoluta - CL", coluna_saida="D", tipo_match="exact")
        rho_cel.value = valor
        rho_cel.api.Locked = True
    
    valor = prop_rho.get("incerteza")
    if valor is not None:
        rho_inc_cel = encontrar_celula(ws, "Densidade absoluta - CL (Incert)", coluna_busca='F', coluna_saida="H", tipo_match="exact")
        rho_inc_cel.value = valor
        rho_inc_cel.api.Locked = True

    valor = amplitudes.get("dpt_alta")
    if valor is not None:
        dpt_alta = encontrar_celula(ws, "Pressão Diferencial Alta (High Differential Pressure)", coluna_saida="D")
        dpt_alta.value = valor
        dpt_alta.api.Locked = True

    valor = amplitudes.get("dp_media")
    if valor is not None:
        dpt_media = encontrar_celula(ws, "Pressão Diferencial Média (Avg Differential Pressure)", coluna_saida="D")
        dpt_media.value = valor
        dpt_media.api.Locked = True

    valor = amplitudes.get("dp_baixa")
    if valor is not None:
        dpt_baixa = encontrar_celula(ws, "Pressão Diferencial Baixa (Low Differential Pressure)", coluna_saida="D")
        dpt_baixa.value = valor
        dpt_baixa.api.Locked = True


    valor = incerteza_abs.get("dpt_alta")
    if valor is not None:
        inc_alta = encontrar_celula(ws, "Pressão diferencial de Alta (Incert)", coluna_saida="D", tipo_match="exact")
        inc_alta.value = valor
        inc_alta.api.Locked = True
        
    
    valor = erro_fid.get("dpt_alta")
    if valor is not None:
        fid_alta = encontrar_celula(ws, "(High Differential Pressure) Erro Fiducial (Fiducial Error)", coluna_saida="D", tipo_match="exact")
        fid_alta.value = valor
        fid_alta.api.Locked = True

        
    valor = incerteza_abs.get("dp_media")
    if valor is not None:
        inc_media = encontrar_celula(ws, "Pressão diferencial de Média (Incert)", coluna_saida="D", tipo_match="exact")
        inc_media.value = valor
        inc_media.api.Locked = True
    
    
    valor = erro_fid.get("dp_media")
    if valor is not None:
        fid_medio = encontrar_celula(ws, "(Medium Range Differential Pressure) Fiducial Error", coluna_saida="D", tipo_match="exact")
        fid_medio.value = valor
        fid_medio.api.Locked = True

        
    valor = incerteza_abs.get("dp_baixa")
    if valor is not None:
        inc_baixa = encontrar_celula(ws, "Pressão diferencial de Baixa  (Incert)", coluna_saida="D", tipo_match="exact")
        inc_baixa.value = valor
        inc_baixa.api.Locked = True
        
        
    valor = erro_fid.get("dp_baixa")
    if valor is not None:
        celu_fid_baixa = encontrar_celula(ws, "(Low Range Differential Pressure) Pressure) Fiducial Error", coluna_saida="D", tipo_match="exact")
        celu_fid_baixa.value = valor
        celu_fid_baixa.api.Locked = True
        
    
    
    valor = incerteza_abs.get("pressao_estatica")
    if valor is not None:
        cel_inc_estatica = encontrar_celula(ws, "Pressão estática (Incert)", coluna_saida="D", tipo_match="exact")
        cel_inc_estatica.value = valor
        cel_inc_estatica.api.Locked = True
        
    

    valor = erro_fid.get("pressao_estatica")
    if valor is not None:
        celu_fid_estatica = encontrar_celula(ws, "(Static Pressure) Erro Fiducial (Fiducial Error)", coluna_saida="D", tipo_match="exact")
        celu_fid_estatica.value = valor
        celu_fid_estatica.api.Locked = True
        

    valor_k = k_val.get("dpt_alta")
    if valor_k is not None:
        cel_k_alta = encontrar_celula(ws, "K factor (Alta)", coluna_saida="D", tipo_match="exact")
        cel_k_alta.value=valor_k
        cel_k_alta.api.Locked = True
        
        
    valor_k = k_val.get("dp_media")
    if valor_k is not None:
        cl_k_media = encontrar_celula(ws, "K factor (Média)", coluna_saida="D", tipo_match="exact")
        cl_k_media.value = valor_k
        cl_k_media.api.Locked = True

    valor_k = k_val.get("dp_baixa")
    if valor_k is not None:
        cl_k_baixa = encontrar_celula(ws, "K factor (Baixa)", coluna_saida="D", tipo_match="exact")
        cl_k_baixa.value = valor_k
        cl_k_baixa.api.Locked = True
        
    
    valor_k = k_val.get("pressao_estatica")
    if valor_k is not None:
        cl_k_estatica = encontrar_celula(ws, "K factor estática", coluna_saida="D", tipo_match="exact")
        cl_k_estatica.value = valor_k
        cl_k_estatica.api.Locked = True

    inc_transm = incert_transm.get("incerteza") if incert_transm else None
    k_trasm = incert_transm.get("k") if incert_transm else None
    err_transm = incert_transm.get("erro") if incert_transm else None
    
    if inc_transm is not None:
        cel_inc_transm = encontrar_celula(ws, "Inc transm",coluna_busca='F' ,coluna_saida="H", tipo_match="exact")
        cel_inc_transm.value = inc_transm
        cel_inc_transm.api.Locked = True
        cel_k_transm = encontrar_celula(ws, "k transm",coluna_busca='F' ,coluna_saida="H", tipo_match="exact")
        cel_k_transm.value = k_trasm
        cel_k_transm.api.Locked = True     
        cel_err_transm = encontrar_celula(ws, "erro residual transm",coluna_busca='F' ,coluna_saida="H", tipo_match="exact")
        cel_err_transm.value = err_transm       
        cel_err_transm.api.Locked = True
        
    
    inc_termo = incert_termo.get("incerteza") if incert_termo else None
    k_termo = incert_termo.get("k") if incert_termo else None
    err_termo = incert_termo.get("erro") if incert_termo else None
    
    if inc_termo is not None:
        cel_inc_termo = encontrar_celula(ws, "Inc termo", coluna_saida="D", tipo_match="exact")
        cel_inc_termo.value = inc_termo
        cel_inc_termo.api.Locked = True
        cel_k_termo = encontrar_celula(ws, "k termo", coluna_saida="D", tipo_match="exact")
        cel_k_termo.value = k_termo
        cel_k_termo.api.Locked = True
        cel_err_termo = encontrar_celula(ws, "erro residual termo",coluna_saida="D", tipo_match="exact")
        cel_err_termo.value = err_termo
        cel_err_termo.api.Locked = True
        
    if icert_comb is not None:
        incert_temp = encontrar_celula(ws, "Temperatura (Incert)", coluna_saida="D", tipo_match="exact")
        fid_temp = encontrar_celula(ws, "(Temperature) Erro Fiducial (Fiducial Error)", coluna_saida="D", tipo_match="exact")
        k_temp = encontrar_celula(ws, "K factor Temp", coluna_saida="D", tipo_match="exact")
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
    """
    Preenche a aba "Meter run parameters" com os dados metrológicos da placa de orifício.

    Valores escritos (todos bloqueados contra edição após escrita):
      - Diâmetro do orifício medido (valor médio do certificado).
      - Incerteza expandida do diâmetro.
      - Fator K da placa de orifício.
      - Coeficiente de expansão térmica do material da placa.

    Args:
        wb: Workbook xlwings aberto da planilha de CI.
        dados (dict): Dados consolidados. A chave "placa" deve conter os dados
                      do certificado de calibração da placa de orifício.
    """

    placa_dados = dados_placa(dados)
    ws = wb.sheets["Parameters"]

    diametro = placa_dados.get("diametro_orificio", {}).get("valor", None)
    incert = placa_dados.get("diametro_orificio", {}).get("incerteza", None)
    k_placa = placa_dados.get("diametro_orificio", {}).get('k', None)
    coef_placa = placa_dados.get("coef_dilatacao", None)

    if diametro is not None:
        cel_diamentro_po =encontrar_celula(ws, "Diâmetro do orificio medido (Orifice bore Diameter)", coluna_busca="L", coluna_saida="N", tipo_match="exact")
        cel_diamentro_po.value = diametro
        cel_diamentro_po.api.Locked = True
        
    
    if incert is not None:
        cel_incert_po = encontrar_celula(ws, "(Uncertainty) PO", coluna_busca="L", coluna_saida="N", tipo_match="exact")
        cel_incert_po.value = incert
        cel_incert_po.api.Locked = True
        

    if k_placa is not None:
        k_placa_cel = encontrar_celula(ws, "K factor PO", coluna_busca="L", coluna_saida="N", tipo_match="exact")
        k_placa_cel.value = k_placa
        k_placa_cel.api.Locked = True
    

    if coef_placa is not None:
        coef_placa_cel = encontrar_celula(ws, "Coeficiente de exp. Térmica (Thermal coefficient PO)", coluna_busca="L", coluna_saida="N")
        coef_placa_cel.value = coef_placa
        coef_placa_cel.api.Locked = True
    
    placa_dados= None
    diametro = None
    incert = None
    k_placa = None
    coef_placa = None
     
def preencher_cromatografia(wb, dados):
    """
    Preenche a aba "Chromatography" com a composição do gás e suas propriedades
    extraídas do certificado de análise cromatográfica.

    Comportamento:
      - Limpa o intervalo B2:E200 antes de escrever, evitando dados residuais
        de revisões anteriores.
      - H2S e componentes com "HIDROG" no nome são excluídos da escrita por
        restrição operacional.
      - Escreve em sequência: componentes (rótulo, nome, mol%, incerteza),
        propriedades em condição padrão e propriedades em condições de amostragem.
      - Retorna sem escrever caso não haja dados de cromatografia ou nenhum
        componente válido, preservando o conteúdo anterior da planilha.

    Args:
        wb: Workbook xlwings aberto da planilha de CI.
        dados (dict): Dados consolidados. A chave "cromatografia" deve conter
                      "componentes", "propriedades_condicao_padrao" e
                      "propriedades_condicoes_amostragem".
    """
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

    propriedades_padrao = cromatografia.get("propriedades_condicao_padrao", {})

    for nome, prop in propriedades_padrao.items():
        ws.range(f"B{linha}").value = nome
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

    propriedades_amostragem = cromatografia.get("propriedades_condicoes_amostragem", {})

    for nome, prop in propriedades_amostragem.items():
        ws.range(f"B{linha}").value = nome
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
    """
    Preenche a aba "Equipment List" com os dados de identificação de cada instrumento
    (TAG, número de série e certificado de calibração).

    A linha de cada instrumento é localizada dinamicamente pelo texto da coluna A,
    eliminando dependência de posições fixas e tornando a função resiliente a
    variações de layout na planilha.

    Instrumentos cobertos: placa de orifício, transmissor de temperatura,
    termoresistência, pressão estática, DPT Alta, DPT Média, DPT Baixa
    e cromatografia (somente número de certificado, coluna F).

    Args:
        wb: Workbook xlwings aberto da planilha de CI.
        dados (dict): Dados consolidados contendo os sub-dicts de cada instrumento
                      e o dict "cromatografia" com seu cabeçalho.
    """

    sec_dados = dados_secundários(dados)
    placa_dados = dados_placa(dados)

    ws = wb.sheets["Equipment List"]

    instrumentos = {
        "placa":            "Placa de Orifício (Orifice Plate)",
        "temperatura":      "Transmissor de Temperatura (Temperature Transmitter)",
        "termoresistencia": "Termorresistência (Thermoresistance)",
        "pressao_estatica": "Pressão Estática (Static Pressure)",
        "dpt_alta":         "Pressão Diferencial Alta (High Differential Pressure)",
        "dp_baixa":         "Pressão Diferencial Baixa (Low Differential Pressure)",
        "dp_media":         "Pressão Diferencial Média (Avg Differential Pressure)",
        'cromatografia':    "Cromatografia (Gas Chromatography)"
    }

    for chave, texto in instrumentos.items():
        cel = encontrar_celula(ws, texto, coluna_busca="A", coluna_saida="A", tipo_match="contains")
        if cel is None:
            continue

        linha = cel.row

        if chave == "cromatografia":
            cert = dados.get("cromatografia", {}).get("cabecalho", {}).get("certificado")
            if cert is not None:
                ws.range(f"F{linha}").value = cert
            continue

        info = placa_dados if chave == "placa" else sec_dados.get(chave, {})

        tag  = info.get("tag")
        ns   = info.get("numero_serie")
        cert = info.get("certificado")

        if tag  is not None: ws.range(f"D{linha}").value = tag
        if ns   is not None: ws.range(f"E{linha}").value = ns
        if cert is not None: ws.range(f"F{linha}").value = cert

    sec_dados = None
    placa_dados = None

def preencher_report(wb, dados):
    """
    Preenche a aba "Report" com o número de cálculo atualizado e o motivo da revisão
    gerado automaticamente com base nos tipos de dados importados pelo usuário.

    Lógica do motivo de revisão (baseada nas respostas de importação XML):
      - Placa + Cromatografia + Secundários → texto composto completo
      - Qualquer combinação de dois tipos     → texto combinado
      - Tipo único                            → texto específico
      - Nenhum dado importado                 → string vazia

    O número de cálculo é incrementado diretamente no texto da célula de título
    via utilitário alterar_ncalculo(), refletindo a nova revisão gerada.

    Args:
        wb: Workbook xlwings aberto da planilha de CI.
        dados (dict): Dados consolidados (não utilizado diretamente; as respostas
                      de importação são obtidas via obter_respostas()).
    """

    respostas = obter_respostas()
    ws = wb.sheets["Report"]

    celula_titulo = encontrar_celula(ws,"Relatório de Cálculo de Incerteza",coluna_busca="B",  coluna_saida="B", tipo_match="contains")
    if celula_titulo is not None:
        texto_atual = celula_titulo.value

        if texto_atual:
            novo_texto = alterar_ncalculo(texto_atual)
            celula_titulo.value = novo_texto
        else:
            print("Célula encontrada, mas vazia")
    else:
        print("Célula não encontrada")

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

    motivo_ci = encontrar_celula(ws, "Motivo da Revisão (Reason for Revision)", coluna_busca="B",coluna_saida="C", tipo_match="exact")
    motivo_ci.clear_contents()
    motivo_ci.value = texto



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

        for ws in wb.sheets:
            ws.api.Unprotect()

        preencher_gas_parameters(wb, dados)
        preencher_meter_run_parameter(wb, dados)
        preencher_cromatografia(wb, dados)
        preencher_equipament_list(wb, dados)
        preencher_report(wb, dados)
        #app.api.Run("AUTOMATICO")
        #app.api.Visible = False  # macro pode sobrescrever visible=False; força de volta

        #app.calculate()

        wb.save()

        wb.close()

    finally:

        app.quit()


