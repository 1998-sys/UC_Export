def faixas_calibradas(dados): 
    instrumentos_alvo = [
        "dpt_alta",
        "dp_media",
        "dp_baixa",
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


def validar_ordem_dpts(dados):

    faixas = faixas_calibradas(dados)
    amplitudes = calcular_amplitudes(faixas)

    dpts = {
        "dp_baixa": amplitudes.get("dp_baixa"),
        "dp_media": amplitudes.get("dp_media"),
        "dpt_alta": amplitudes.get("dpt_alta"),
    }

    dpts_validos = {k: v for k, v in dpts.items() if v is not None}

    quantidade = len(dpts_validos)


    if quantidade <= 1:
        return True, None

    valores = list(dpts_validos.values())

    
    if len(set(valores)) == 1:
        return False, (
            "Os DPTs carregados possuem o mesmo range.\n\n"
            "Isso pode indicar que o mesmo certificado foi "
            "selecionado mais de uma vez.\n\n"
            "Por favor, revise os arquivos."
        )

    ordenado = sorted(dpts_validos.items(), key=lambda x: x[1])
    ordem_atual = [item[0] for item in ordenado]

    ordem_logica = ["dp_baixa", "dp_media", "dpt_alta"]

    ordem_esperada = [d for d in ordem_logica if d in dpts_validos]

    if ordem_atual != ordem_esperada:
        return False, (
            "A ordem dos ranges dos DPTs está incorreta.\n\n"
            f"Ordem encontrada: {ordem_atual}\n"
            f"Ordem esperada: {ordem_esperada}\n\n"
            "Os dados podem ter sido carregados incorretamente."
        )

    return True, None



