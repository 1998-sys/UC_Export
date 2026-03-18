def incerteza_temp_oleo(dados):
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

    for ponto in pontos:
        inc = ponto.get("incerteza")

        if inc in (None, "NI", ""):
            continue

        try:
            inc_val = float(inc)
            erro_val = abs(float(ponto.get("erro", 0)))
        except (ValueError, TypeError):
            continue

        # maior erro
        if maior_erro is None or erro_val > maior_erro:
            maior_erro = erro_val

        # maior incerteza
        if maior_incerteza is None or inc_val > maior_incerteza:
            maior_incerteza = inc_val

    if maior_erro is None and maior_incerteza is None:
        return None

    return {
        "maior_incerteza": maior_incerteza,
        "maior_erro": maior_erro
    }


def incerteza_percentual(dados, amplitudes):
    resultado = {}

    for nome in amplitudes.keys():
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

            resultado[nome] = float(incerteza_percentual)

        except (ValueError, TypeError):
            resultado[nome] = None

    return resultado


def erro_fiducial(dados, amplitudes):
    resultado = {}

    for nome in amplitudes.keys():
        instrumento = dados.get(nome)

        if not instrumento:
            resultado[nome] = None
            continue

        try:
            erro_fiducial = (
                instrumento
                .get("instrumento", {})
                .get("as_found", {})
                .get("erro_fiducial")
            )

            if erro_fiducial is None:
                resultado[nome] = None
                continue

            resultado[nome] = float(erro_fiducial)

        except (ValueError, TypeError):
            resultado[nome] = None

    return resultado