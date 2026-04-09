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


