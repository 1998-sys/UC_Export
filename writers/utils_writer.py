import math



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
            incerteza_absoluta = (amplitude * incerteza_percentual)/100
            resultado[nome] = round(incerteza_absoluta, 6)

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
            erro_fiducial_absoluto = (amplitude * erro_fiducial_percentual)/100
            resultado[nome] = round(erro_fiducial_absoluto, 6)

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

    u_transm = float(incert_transm.get("incerteza", 0))
    u_termo = float(incert_termo.get("incerteza", 0))

    u_combinada = math.sqrt(u_transm**2 + u_termo**2)

    erro_transm = float(incert_transm.get("erro"))
    erro_termo = float(incert_termo.get("erro"))

    erro_combinado = math.sqrt(erro_transm**2 + erro_termo**2)

    return {
        "incerteza": round(u_combinada, 6),
        "erro": round(erro_combinado, 6),
        "k": 2
    }

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

