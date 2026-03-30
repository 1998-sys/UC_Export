import math
from writers.util_writer_oleo import normalizar

def faixas_calibradas(dados): 
    instrumentos_alvo = [
        "dpt_alta",
        "dp_media",
        "dp_baixa",
        "pressao_estatica",
    ]

    resultado = {}
    for nome in instrumentos_alvo:
        instrumento = dados.get(nome)

        if instrumento is None:
            resultado[nome] = {"min": None, "max": None}
            continue

        faixa = (
            instrumento
            .get("instrumento", {})
            .get("as_found", {})
            .get("faixa_calibrada", {})
        )

        resultado[nome] = {
            "min": faixa.get("min"),
            "max": faixa.get("max")
        }
    return resultado

def calcular_amplitudes(faixas):
    resultado = {}

    for nome, valores in faixas.items():
        min_val = valores.get("min")
        max_val = valores.get("max")

        try:
            if min_val is None or max_val is None:
                resultado[nome] = None
            else:
                resultado[nome] = float(max_val) - float(min_val)
        except (ValueError, TypeError):
            resultado[nome] = None

    return resultado

def formatar_celula_valor(cell):
    cell.api.Font.Name = "Calibri"
    cell.api.Font.Size = 10
    cell.color = (255, 204, 153) 

def incerteza_absoluta(dados, amplitudes):
    resultado = {}
    for nome, amplitude in amplitudes.items():
        if amplitude is None:
            resultado[nome] = None
            continue
        instrumento = dados.get(nome)
        if not instrumento:
            resultado[nome] = None
            continue
        try:
            incerteza_percentual = (
                instrumento
                .get("instrumento", {})
                .get("as_found", {})
                .get("incerteza_percentual")
            )

            if incerteza_percentual is None:
                resultado[nome] = None
                continue
            incerteza_percentual = float(incerteza_percentual)
            formula = f"={amplitude}*{incerteza_percentual}%"
            resultado[nome] = formula

        except (ValueError, TypeError):
            resultado[nome] = None

    return resultado

def erro_fiducial_abs(dados, amplitudes):
    resultado = {}
    for nome, amplitude in amplitudes.items():
        if amplitude is None:
            resultado[nome] = None
            continue
        instrumento = dados.get(nome)
        if not instrumento:
            resultado[nome] = None
            continue
        try:
            erro_fiducial_percentual = (
                instrumento
                .get("instrumento", {})
                .get("as_found", {})
                .get("erro_fiducial")
            )

            if erro_fiducial_percentual is None:
                resultado[nome] = None
                continue

            erro_fiducial_percentual = float(erro_fiducial_percentual)
            formula = f"={amplitude}*{erro_fiducial_percentual}%"

            resultado[nome] = formula

        except (ValueError, TypeError):
            resultado[nome] = None
    
    return resultado

def obter_k(dados):
    resultado = {}

    for nome, instrumento_data in dados.items():

        if instrumento_data.get("tipo") != "pressao":
            continue

        pontos = (
            instrumento_data
            .get("instrumento", {})
            .get("as_found", {})
            .get("pontos", [])
        )

        maior_incerteza = None
        ks_mesma_incerteza = []

        for ponto in pontos:

            inc = ponto.get("incerteza")
            k_val = ponto.get("k")

            # Ignorar NI
            if inc in (None, "NI") or k_val in (None, "NI"):
                continue

            try:
                inc = float(inc)
                k_val = float(k_val)
            except ValueError:
                continue

            if maior_incerteza is None:
                maior_incerteza = inc
                ks_mesma_incerteza = [k_val]

            elif inc > maior_incerteza:
                maior_incerteza = inc
                ks_mesma_incerteza = [k_val]

            elif inc == maior_incerteza:
                ks_mesma_incerteza.append(k_val)

        if maior_incerteza is None:
            resultado[nome] = None
        else:
            resultado[nome] = max(ks_mesma_incerteza)

    return resultado

def incerteza_temperatura(dados):
    if not dados:
        return None

    pontos = (
        dados
        .get("instrumento", {})
        .get("as_found", {})
        .get("pontos", [])
    )

    maior_erro = None
    maior_incerteza = None
    candidatos = []

    for ponto in pontos:
        inc = ponto.get("incerteza")

        if inc in (None, "NI", ""):
            continue

        try:
            inc_val = float(inc)
            k_val = float(ponto.get("k", 0))
            erro_val = abs(float(ponto.get("erro", 0)))
        except (ValueError, TypeError):
            continue

    
        if maior_erro is None or erro_val > maior_erro:
            maior_erro = erro_val

        # encontra maior incerteza
        if maior_incerteza is None or inc_val > maior_incerteza:
            maior_incerteza = inc_val
            candidatos = [(ponto, k_val)]
        elif inc_val == maior_incerteza:
            candidatos.append((ponto, k_val))

    if not candidatos:
        return None

    # desempate pelo maior k
    melhor_ponto = max(candidatos, key=lambda x: x[1])[0]

    # substitui erro pelo maior erro global
    melhor_ponto["erro"] = maior_erro

    return melhor_ponto

def incert_temp_comb(incert_transm, incert_termo):

    if not incert_termo:
        return incert_transm

    if not incert_transm:
        return incert_termo

    try:
        u_transm = incert_transm.get("incerteza")
        u_termo = incert_termo.get("incerteza")

        erro_transm = incert_transm.get("erro")
        erro_termo = incert_termo.get("erro")

        if None in (u_transm, u_termo, erro_transm, erro_termo):
            return None

        formula_incerteza = f"=SQRT(({u_transm})^2 + ({u_termo})^2)"
        formula_erro = f"=SQRT(({erro_transm})^2 + ({erro_termo})^2)"

        return {
            "incerteza": formula_incerteza,
            "erro": formula_erro,
            "k": 2
        }

    except (ValueError, TypeError):
        return None

def dados_secundários(dados):

    instrumentos_alvo = [
        "dpt_alta",
        "dp_media",
        "dp_baixa",
        "pressao_estatica",
        "temperatura",
        "termoresistencia",
    ]

    resultado = {}

    for nome in instrumentos_alvo:
        instrumento = dados.get(nome)

        if instrumento is None:
            resultado[nome] = {
                "certificado": None,
                "numero_serie": None,
                "tag": None
            }
            continue

        dados_instrumento = instrumento.get("instrumento", {})

        if nome == "temperatura":
            transmissor = dados_instrumento.get("transmissor", {})

            numero_serie = transmissor.get("numero_serie")
            tag = transmissor.get("tag")

        else:
            numero_serie = dados_instrumento.get("numero_serie")
            tag = dados_instrumento.get("tag")

        resultado[nome] = {
            "certificado": instrumento.get("numero_certificado"),
            "numero_serie": numero_serie,
            "tag": tag,
        }

    return resultado

def dados_placa(dados):

    placa = dados.get("placa")

    if placa is None:
        return {
            "certificado": None,
            "numero_serie": None,
            "tag": None,
            "coef_dilatacao": None,
            "diametro_orificio": {
                "valor": None,
                "incerteza": None,
                "k": None,
                "aprovado": None
            }
        }

    dados_placa = placa.get("placa", {})

    numero_serie = dados_placa.get("numero_serie")
    tag = dados_placa.get("tag")
    coef = dados_placa.get("coef_dilatacao")

    diametro = dados_placa.get("diametro_orificio", {}).get("valor_medio", {})

    return {
        "certificado": placa.get("numero_certificado"),
        "numero_serie": numero_serie,
        "tag": tag,
        "coef_dilatacao": coef,
        "diametro_orificio": {
            "valor": diametro.get("valor"),
            "incerteza": diametro.get("incerteza"),
            "k": diametro.get("k"),
            "aprovado": diametro.get("aprovado")
        }
    }

respostas_xml = {}

def registrar_resposta(chave, valor):
    respostas_xml[chave] = valor

def obter_respostas():
    return respostas_xml


def encontrar_celula_pressao_ref(ws):
    texto_ref = normalizar(
        "Pressão estática (static pressure), P"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")

    print('célula da pressão de referencia não encontrada')
    return None

def encontrar_celula_temperatura_ref(ws):
    texto_ref = normalizar(
        "Temperatura (Temperature), T"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")

    print('Célula da Temperatura de Referencia não encotrada')
    return None

def celula_pressao_dif_alta(ws):
    texto_ref = normalizar(
        "Pressão Diferencial Alta (High Differential Pressure)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")

    print('Pressão diferencial alta não encontrada')
    return None

def celula_pressao_dif_media(ws):
    texto_ref = normalizar(
        "Pressão Diferencial Média (Avg Differential Pressure)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")

    print('Pressão diferencial média não encontrada')
    return None

def celula_pressao_dif_baixa(ws):
    texto_ref = normalizar(
        "Pressão Diferencial Baixa (Low Differential Pressure)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")

    print('Pressão diferencial baixa não encontrada')
    return None

def celula_incerteza_alta(ws):
    texto_ref = normalizar(
        "Pressão diferencial de Alta"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref == valor:
            return ws.range(f"E{i}")

    print('Pressão diferencial não encontrada')
    return None

def celula_fid_alta(ws):
    texto_ref = normalizar(
        "(High Differential Pressure) Erro Fiducial (Fiducial Error)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref == valor:
            return ws.range(f"E{i}")

    print('Pressão diferencial não encontrada')
    return None

def celula_incerteza_media(ws):
    texto_ref = normalizar(
        "Pressão diferencial de Média"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
            return ws.range(f"E{i}")

    print('Pressão diferencial de média não encontrada')
    return None

def celula_fid_media(ws):
    texto_ref = normalizar(
        "(Medium Range Differential Pressure) Fiducial Error"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula de erro fiducial de média {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")

    print('Pressão diferencial de média não encontrada')
    return None

def celula_incert_baixa(ws):
    texto_ref = normalizar(
        "Pressão diferencial de Baixa"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula de incerteza de baixa {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")

    print('Pressão diferencial de baixa não encontrada')
    return None

def celula_fid_baixa(ws):
    texto_ref = normalizar(
        "(Low Range Differential Pressure) Pressure) Fiducial Error"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
           # print(f"Encontrado célula de erro fiducial de baixa {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")

    print('erro fiducial de baixa não encontrada')
    return None


def celula_inc_estatica(ws):
    texto_ref = normalizar(
        "Pressão estática"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula de erro fiducial de baixa {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")

    print('erro fiducial de baixa não encontrada')
    return None

def celula_fid_estatica(ws):
    texto_ref = normalizar(
        "(Static Pressure) Erro Fiducial (Fiducial Error)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula de erro fiducial de baixa {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")

    print('erro fiducial de baixa não encontrada')
    return None

def celula_k_alta(ws):
    texto_ref = normalizar(
        "K factor (Alta)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"G{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula k de alta {i+2}: {ws.range(f'E{i}').value}")
            return ws.range(f"G{i+2}")

    print('K factor (Alta) não encontrado')
    return None

def celula_k_media(ws):
    texto_ref = normalizar(
        "K factor (Média)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"G{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula k de alta {i+2}: {ws.range(f'E{i}').value}")
            return ws.range(f"G{i+2}")

    print('K factor (Média) não encontrado')
    return None


def celula_k_baixa(ws):
    texto_ref = normalizar(
        "K factor (Baixa)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"G{i}").value)

        if valor and normalizar(valor) == texto_ref:
            #print(f"Encontrado célula k de alta {i+2}: {ws.range(f'E{i}').value}")
            return ws.range(f"G{i+2}")

    print('K factor (Baixa) não encontrado')
    return None


def celula_k_estatica(ws):
    texto_ref = normalizar(
        "K factor estática"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"G{i}").value)

        if valor and normalizar(valor) == texto_ref:
            print(f"Encontrado célula k de alta {i+2}: {ws.range(f'E{i}').value}")
            return ws.range(f"F{i}")

    print('K factor presão estática não encontrado')
    return None


def celula_inc_temp(ws):
    texto_ref = normalizar(
        "Temperatura"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref == valor:
            print(f"Encontrado célula incerteza de temperatura {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"F{i}")

    print('célula da incerteza de temperatura não encontrada')
    return None



def celula_fid_temp(ws):
    texto_ref = normalizar(
        "(Temperature) Erro Fiducial (Fiducial Error)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref == valor:
            print(f"Encontrado célula incerteza de temperatura {i}: {ws.range(f'E{i}').value}")
            return ws.range(f"F{i}")

    print('célula da erro fiducial de temperatura não encontrada')
    return None




