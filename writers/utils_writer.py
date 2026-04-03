import math
from writers.util_writer_oleo import normalizar
import os
import re

def faixas_calibradas(dados):
    """
    Extrai a faixa calibrada (min/max) de cada instrumento de pressão dos dados coletados.

    Parâmetros:
        dados (dict): Dicionário com os dados coletados dos instrumentos.
                      Chaves esperadas: "dpt_alta", "dp_media", "dp_baixa", "pressao_estatica".

    Retorno:
        dict: Mapeamento { nome_instrumento: { "min": float|None, "max": float|None } }
    """
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
    """
    Calcula a amplitude (max - min) de cada instrumento a partir das faixas calibradas.

    Parâmetros:
        faixas (dict): Resultado de faixas_calibradas().

    Retorno:
        dict: Mapeamento { nome_instrumento: float|None }
              None quando min ou max estiver ausente ou inválido.
    """
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
    """
    Aplica formatação padrão a uma célula de valor preenchido pelo sistema:
    fonte Calibri 10pt e fundo laranja claro (255, 204, 153).

    Parâmetros:
        cell: Objeto de célula xlwings.
    """
    cell.api.Font.Name = "Calibri"
    cell.api.Font.Size = 10
    cell.color = (255, 204, 153) 

def incerteza_absoluta(dados, amplitudes):
    """
    Gera fórmulas Excel de incerteza absoluta para cada instrumento.
    Fórmula: amplitude * incerteza_percentual%

    Parâmetros:
        dados (dict): Dados coletados dos instrumentos.
        amplitudes (dict): Resultado de calcular_amplitudes().

    Retorno:
        dict: Mapeamento { nome_instrumento: str|None }
              Valor é uma string de fórmula Excel (ex: "=150.0*0.075%") ou None.
    """
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
    """
    Gera fórmulas Excel de erro fiducial absoluto para cada instrumento.
    Fórmula: amplitude * erro_fiducial_percentual%

    Parâmetros:
        dados (dict): Dados coletados dos instrumentos.
        amplitudes (dict): Resultado de calcular_amplitudes().

    Retorno:
        dict: Mapeamento { nome_instrumento: str|None }
              Valor é uma string de fórmula Excel ou None.
    """
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
    """
    Determina o fator K de cobertura para cada instrumento de pressão.
    Seleciona o K associado ao ponto de maior incerteza; em caso de empate,
    retorna o maior K entre os candidatos. Pontos com valor "NI" são ignorados.

    Parâmetros:
        dados (dict): Dados coletados dos instrumentos. Apenas instrumentos
                      com tipo == "pressao" são processados.

    Retorno:
        dict: Mapeamento { nome_instrumento: float|None }
    """
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
    """
    Identifica o ponto crítico de calibração de um instrumento de temperatura.
    Critério: maior incerteza absoluta; desempate pelo maior K.
    O campo "erro" do ponto retornado é substituído pelo maior erro absoluto
    encontrado em todos os pontos do instrumento.

    Parâmetros:
        dados (dict): Dados de um único instrumento (transmissor ou termoresistência).

    Retorno:
        dict|None: Ponto de calibração crítico com chaves "incerteza", "erro", "k",
                   ou None se não houver pontos válidos.
    """
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
    """
    Combina as incertezas do transmissor de temperatura e da termoresistência
    pela regra da raiz da soma dos quadrados (RSS), gerando fórmulas Excel.
    Se apenas um dos instrumentos estiver disponível, retorna o existente sem combinação.

    Parâmetros:
        incert_transm (dict|None): Ponto crítico do transmissor (saída de incerteza_temperatura).
        incert_termo (dict|None): Ponto crítico da termoresistência (saída de incerteza_temperatura).

    Retorno:
        dict|None: { "incerteza": str (fórmula Excel), "erro": str (fórmula Excel), "k": 2 }
                   ou None em caso de dados insuficientes.
    """
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

        formula_incerteza = f"=SQRT(SUMSQ({u_transm},{u_termo}))"
        formula_erro = f"=SQRT(SUMSQ({erro_transm},{erro_termo}))"

        return {
            "incerteza": formula_incerteza,
            "erro": formula_erro,
            "k": 2
        }

    except (ValueError, TypeError):
        return None

def dados_secundários(dados):
    """
    Extrai os dados de identificação (certificado, número de série, tag) dos instrumentos
    secundários: DPTs (alta/média/baixa), pressão estática, transmissor de temperatura
    e termoresistência.

    Parâmetros:
        dados (dict): Dados coletados dos instrumentos.

    Retorno:
        dict: Mapeamento { nome_instrumento: { "certificado", "numero_serie", "tag" } }
              Para o transmissor de temperatura, os campos são extraídos do sub-dict "transmissor".
    """
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
    """
    Extrai os dados de identificação e metrológicos da placa de orifício.

    Parâmetros:
        dados (dict): Dados coletados. Chave esperada: "placa".

    Retorno:
        dict: { "certificado", "numero_serie", "tag", "coef_dilatacao",
                "diametro_orificio": { "valor", "incerteza", "k", "aprovado" } }
              Todos os campos são None se a placa não estiver presente.
    """
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
    """
    Registra a resposta do usuário para uma etapa de importação XML.

    Parâmetros:
        chave (str): Identificador da etapa (ex: "dpt_alta", "placa").
        valor (bool): True se o usuário confirmou a importação, False caso contrário.
    """
    respostas_xml[chave] = valor

def obter_respostas():
    """
    Retorna o dicionário com todas as respostas de importação XML registradas na sessão.

    Retorno:
        dict: Mapeamento { chave: bool }
    """
    return respostas_xml

def encontrar_celula(
    ws,
    texto_busca,
    coluna_busca="B",
    coluna_saida="F",
    tipo_match="contains",
    offset_linha=0,
    debug=False
):
    """
    Localiza uma célula na planilha pelo conteúdo textual de uma coluna de referência
    e retorna a célula de destino na mesma linha (com offset opcional).

    Parâmetros:
        ws: Sheet xlwings da planilha ativa.
        texto_busca (str): Texto a ser buscado (será normalizado internamente).
        coluna_busca (str): Coluna onde o texto será procurado. Padrão: "B".
        coluna_saida (str): Coluna da célula a ser retornada. Padrão: "F".
        tipo_match (str): Modo de comparação:
                          "contains" — texto_busca contido no valor da célula (padrão);
                          "exact"    — correspondência exata.
        offset_linha (int): Deslocamento de linhas aplicado à linha encontrada antes de
                            retornar a célula de saída. Padrão: 0.
        debug (bool): Se True, imprime no console a linha encontrada ou a ausência do texto.

    Retorno:
        xlwings.Range | None: Célula de destino, ou None se o texto não for encontrado.
    """
    texto_ref = normalizar(texto_busca)

    last_row = ws.range(coluna_busca + str(ws.cells.last_cell.row)).end("up").row

    for i in range(1, last_row + 1):
        valor_celula = ws.range(f"{coluna_busca}{i}").value
        valor = normalizar(valor_celula)

        if not valor:
            continue

        if (
            (tipo_match == "contains" and texto_ref in valor)
            or (tipo_match == "exact" and texto_ref == valor)
        ):
            linha_saida = i + offset_linha

            if debug:
                print(f"Encontrado '{texto_busca}' na linha {i} → retorno {coluna_saida}{linha_saida}")

            return ws.range(f"{coluna_saida}{linha_saida}")

    if debug:
        print(f"'{texto_busca}' não encontrado")

    return None

def incrementar_nome(caminho_excel):
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

def alterar_ncalculo(texto):
    def repl(match):
        trecho = match.group(0)

        match_num = re.search(r'(\d+)$', trecho)
        if match_num:
            numero = match_num.group(1)
            novo_numero = str(int(numero) + 1).zfill(len(numero))
            return trecho[:match_num.start()] + novo_numero

        return trecho

    # pega tudo que começa com -UCG- até o final do código
    return re.sub(r'-UCG-[\w-]+', repl, texto)

