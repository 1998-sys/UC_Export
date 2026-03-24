import unicodedata

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

def formatar_percentual(valor):
    if valor is None:
        return None

    valor = str(valor).strip()

    if valor == "":
        return None

    if "%" in valor:
        return valor

    return f"{valor}%"

def normalizar(texto):
    if not texto:
        return ""
    texto = str(texto).lower().strip().replace("\n", " ")
    return ''.join(
        c for c in unicodedata.normalize('NFD', texto)
        if unicodedata.category(c) != 'Mn'
    )

def encontrar_celula_resolucao(ws):
    texto_ref = normalizar(
        "Resolução da Termoresistência (Termoresistance resolution)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"M{i}")

    return None

def encontrar_celula_erro_fiducial(ws):
    texto_ref = normalizar(
        "C2.2.2 - Erro Fiducial (Fiducial Error)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"E{i}")

    return None

def encontrar_celula_incerteza_pressao(ws):
    texto_ref = normalizar(
        "Incerteza da calibração do medidor de pressão (Pressure meter calibration uncertainty)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:  
            return ws.range(f"E{i}")
    return None

def encontrar_celula_erro_fiducial_pressao(ws):
    texto_ref = normalizar(
        "C3.1.2 - Erro Fiducial (Fiducial Error)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"E{i}")
    return None

def encontrar_celula_bsw_maximo(ws):
    texto_ref = normalizar(
        "BSW Máximo  (Max BSW Allowed)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            print(f"Encontrado texto de referência na célula B{i}: {ws.range(f'F{i}').value}")
            return ws.range(f"F{i}")
    print('BSW máximo não encontrado')
    return None

def encontrar_celula_incerteza_bsw(ws):
    texto_ref = normalizar(
        "C5.1 Incerteza padrão combinada - BSW (BSW Combined Uncertainty)"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            print(f"Encontrado texto de referência na célula B{i}: {ws.range(f'E{i}').value}")
            return ws.range(f"E{i}")
    print('Incerteza BSW não encontrada')

    return None

def encontrar_celula_pressao_op(ws):
    texto_ref = normalizar(
        "Pressão estática (static pressure), P"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            return ws.range(f"F{i}")
    
    return None

def encontrar_celula_densidade_op(ws):
    texto_ref = normalizar(
        "Densidade nas condições De Referência (Standard Density), ρ"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            print(f"Encontrado texto de referência na célula de desindade {i}: {ws.range(f'F{i}').value}")
            return ws.range(f"F{i}")
    print('Densidade de operação não encontrada')
    return None

def encontrar_celula_temp_op(ws):
    texto_ref = normalizar(
        "Temp. da Termoresistência (Termoresistance temp.) - Ta"
    )

    last_row = ws.range("B" + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor = normalizar(ws.range(f"B{i}").value)

        if valor and texto_ref in valor:
            print(f"Encontrado texto de referência na célula B{i}: {ws.range(f'F{i}').value}")
            return ws.range(f"F{i}")
    print('Temperatura de operação não encontrada')
    return None
