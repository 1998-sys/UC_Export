import xml.etree.ElementTree as ET


def identificar_tipo_xml(caminho_xml: str) -> str:
    """
    Identifica o tipo de instrumento de um XML Petrobras pelo root tag.

    Retorno:
        "placa"             → CERTIFICADO_INSPECAO_PLACA_ORIFICIO
        "cromatografia"     → CROMATOGRAFIA
        "pressao"           → CERTIFICADO_CALIBRACAO_PRESSAO
        "temperatura"       → CERTIFICADO_CALIBRACAO_TEMPERATURA
        "termorresistencia" → CERTIFICADO_CALIBRACAO_TEMPERATURA_TE
        "desconhecido"      → qualquer outro root tag
    """
    root = ET.parse(caminho_xml).getroot()
    tag = root.tag

    if "PLACA_ORIFICIO" in tag:
        return "placa"
    elif "CROMATOGRAFIA" in tag:
        return "cromatografia"
    elif "TEMPERATURA_TE" in tag:
        return "termorresistencia"
    elif "TEMPERATURA" in tag:
        return "temperatura"
    elif "PRESSAO" in tag:
        return "pressao"
    return "desconhecido"


def dados_secundarios(caminho_xml: str) -> dict:
    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    ns = {"cal": "http://Petrobras/Medicao/Calibracao"}

    def get_text(elemento, tag):
        if elemento is None:
            return None
        el = elemento.find(tag, ns)
        return el.text.strip() if el is not None and el.text else None

    dados = {}

    tag_raiz = root.tag

    if "PRESSAO" in tag_raiz:
        dados["tipo"] = "pressao"

    elif "TEMPERATURA_TE" in tag_raiz:
        dados["tipo"] = "termorresistencia"

    elif "TEMPERATURA" in tag_raiz:
        dados["tipo"] = "temperatura"

    else:
        dados["tipo"] = "desconhecido"

    
    dados["numero_certificado"] = get_text(root, "NUMERO_CERTIFICADO")
    
    if dados["tipo"] == "pressao":
        instrumento = root.find("INSTRUMENTO_PRESSAO", ns)
        faixa = instrumento.find("FAIXA_NOMINAL", ns) if instrumento is not None else None
        as_found = instrumento.find("CALIBRACAO_AS_FOUND", ns) if instrumento is not None else None

      
        faixa_calibrada = as_found.find("FAIXA_CALIBRADA", ns) if as_found is not None else None
        min_el = faixa_calibrada.find("MIN", ns) if faixa_calibrada is not None else None
        max_el = faixa_calibrada.find("MAX", ns) if faixa_calibrada is not None else None

        dados["instrumento"] = {
            "numero_serie": get_text(instrumento, "NUM_SERIE"),
            "tag": get_text(instrumento, "TAG"),
            "descricao": get_text(instrumento, "DESCRICAO"),
            "fabricante": get_text(instrumento, "FABRICANTE"),
            "modelo": get_text(instrumento, "MODELO"),
            "data_calibracao": get_text(instrumento, "DATA_CALIBRACAO"),
            "faixa_nominal": {
                "min": get_text(faixa, "MIN") if faixa is not None else None,
                "max": get_text(faixa, "MAX") if faixa is not None else None,
            },
            "as_found": {
                "faixa_calibrada": {
                    "min": min_el.text.strip() if min_el is not None and min_el.text else None,
                    "max": max_el.text.strip() if max_el is not None and max_el.text else None,
                    "unidade": min_el.get("UNIDADE_ENG") if min_el is not None else None
                },
                "erro_fiducial": get_text(as_found, "ERRO_FIDUCIAL"),
                "incerteza_percentual": get_text(as_found, "INCERTEZA"),
                "histerese": get_text(as_found, "HISTERESE"),
                "repetibilidade": get_text(as_found, "REPETIBILIDADE"),
                "pontos": []
            }
        }

        if as_found is not None:
            for ponto in as_found.findall("PONTOS_DE_CALIBRACAO/PONTO_DE_CALIBRACAO", ns):
                ciclo1 = ponto.find("CICLO_1", ns)
                ciclo2 = ponto.find("CICLO_2", ns)

                inc_el = ponto.find("INCERTEZA", ns)

                dados["instrumento"]["as_found"]["pontos"].append({
                    "valor_referencia": get_text(ponto, "VALOR_REFERENCIA"),
                    "ciclo_1_asc": get_text(ciclo1, "VALOR_INDICADO_ASCENDENTE"),
                    "ciclo_1_desc": get_text(ciclo1, "VALOR_INDICADO_DESCENDENTE"),
                    "ciclo_2_asc": get_text(ciclo2, "VALOR_INDICADO_ASCENDENTE"),
                    "ciclo_2_desc": get_text(ciclo2, "VALOR_INDICADO_DESCENDENTE"),

                    "incerteza": inc_el.text.strip() if inc_el is not None and inc_el.text else None,
                    "k": inc_el.get("K") if inc_el is not None else None,
                    "grau_liberdade": inc_el.get("GRAU_LIBERDADE") if inc_el is not None else None,

                    "erro": get_text(ponto, "ERRO"),
                })

    
    elif dados["tipo"] == "temperatura":
        instrumento = root.find("INSTRUMENTO_TEMPERATURA", ns)
        as_found = instrumento.find("CALIBRACAO_AS_FOUND", ns) if instrumento is not None else None

        faixa_calibrada = as_found.find("FAIXA_CALIBRADA", ns) if as_found is not None else None
        min_el = faixa_calibrada.find("MIN", ns) if faixa_calibrada is not None else None
        max_el = faixa_calibrada.find("MAX", ns) if faixa_calibrada is not None else None

        sensor = instrumento.find("ELEMENTO_SENSOR", ns) if instrumento is not None else None
        transmissor = instrumento.find("TRANSMISSOR", ns) if instrumento is not None else None

        dados["instrumento"] = {
            "sensor": {
                "numero_serie": get_text(sensor, "NUM_SERIE"),
                "tag": get_text(sensor, "TAG"),
                "descricao": get_text(sensor, "DESCRICAO"),
            },
            "transmissor": {
                "numero_serie": get_text(transmissor, "NUM_SERIE"),
                "tag": get_text(transmissor, "TAG"),
                "descricao": get_text(transmissor, "DESCRICAO"),
                "fabricante": get_text(transmissor, "FABRICANTE"),
                "modelo": get_text(transmissor, "MODELO"),
            },
            "data_calibracao": get_text(instrumento, "DATA_CALIBRACAO"),
            "as_found": {
                "faixa_calibrada": {
                    "min": min_el.text.strip() if min_el is not None and min_el.text else None,
                    "max": max_el.text.strip() if max_el is not None and max_el.text else None,
                    "unidade": min_el.get("UNIDADE_ENG") if min_el is not None else None
                },
                "erro_fiducial": get_text(as_found, "ERRO_FIDUCIAL"),
                "incerteza": get_text(as_found, "INCERTEZA"),
                "repetibilidade": get_text(as_found, "REPETIBILIDADE"),
                "pontos": []
            }
        }

        if as_found is not None:
           for ponto in as_found.findall("PONTOS_DE_CALIBRACAO/PONTO_DE_CALIBRACAO", ns):
            valor_indicado = ponto.find("VALOR_INDICADO", ns)
            inc_exp = valor_indicado.find("INCERTEZA_EXP", ns) if valor_indicado is not None else None

            dados["instrumento"]["as_found"]["pontos"].append({
                "valor_referencia": get_text(ponto, "VALOR_REFERENCIA"),
                "valor_indicado": get_text(valor_indicado, "VALOR"),

                "incerteza": inc_exp.text.strip() if inc_exp is not None and inc_exp.text else None,
                "k": inc_exp.get("K") if inc_exp is not None else None,
                "grau_liberdade": inc_exp.get("GRAU_LIBERDADE") if inc_exp is not None else None,

                "erro": get_text(ponto, "ERRO"),
            })

 
    elif dados["tipo"] == "termorresistencia":

        instrumento = root.find("ELEMENTO_SENSOR_TEMPERATURA", ns)

        as_found = instrumento.find("CALIBRACAO_AS_FOUND", ns) if instrumento is not None else None
        faixa_calibrada = as_found.find("FAIXA_CALIBRADA", ns) if as_found is not None else None

        min_el = faixa_calibrada.find("MIN", ns) if faixa_calibrada is not None else None
        max_el = faixa_calibrada.find("MAX", ns) if faixa_calibrada is not None else None

        dados["instrumento"] = {
            "numero_serie": get_text(instrumento, "NUM_SERIE"),
            "tag": get_text(instrumento, "TAG"),
            "descricao": get_text(instrumento, "DESCRICAO"),
            "data_calibracao": get_text(instrumento, "DATA_CALIBRACAO"),
            "as_found": {
                "faixa_calibrada": {
                    "min": min_el.text.strip() if min_el is not None and min_el.text else None,
                    "max": max_el.text.strip() if max_el is not None and max_el.text else None,
                    "unidade": min_el.get("UNIDADE_ENG") if min_el is not None else None
                },
                "erro_fiducial": get_text(as_found, "ERRO_FIDUCIAL"),
                "incerteza": get_text(as_found, "INCERTEZA"),
                "repetibilidade": get_text(as_found, "REPETIBILIDADE"),
                "pontos": []
            }
        }

        if as_found is not None:
            for ponto in as_found.findall("PONTOS_DE_CALIBRACAO/PONTO_DE_CALIBRACAO", ns):

                valor_indicado = ponto.find("VALOR_INDICADO", ns)
                inc_exp = valor_indicado.find("INCERTEZA_EXP", ns) if valor_indicado is not None else None

                dados["instrumento"]["as_found"]["pontos"].append({
                    "valor_referencia": get_text(ponto, "VALOR_REFERENCIA"),
                    "valor_indicado": get_text(valor_indicado, "VALOR"),
                    "incerteza": inc_exp.text.strip() if inc_exp is not None and inc_exp.text else None,
                    "k": inc_exp.get("K") if inc_exp is not None else None,
                    "grau_liberdade": inc_exp.get("GRAU_LIBERDADE") if inc_exp is not None else None,
                    "erro": get_text(ponto, "ERRO"),
                })
    return dados

def extrair_max_pressao(caminho_xml: str) -> float:
    """
    Retorna o valor MAX da FAIXA_NOMINAL de um XML de pressão,
    usado para ranquear e classificar automaticamente os transmissores.
    """
    root = ET.parse(caminho_xml).getroot()
    ns = {"cal": "http://Petrobras/Medicao/Calibracao"}
    instrumento = root.find("INSTRUMENTO_PRESSAO", ns)
    faixa = instrumento.find("FAIXA_NOMINAL", ns) if instrumento is not None else None
    max_el = faixa.find("MAX", ns) if faixa is not None else None
    if max_el is not None and max_el.text:
        try:
            return float(max_el.text.strip().replace(",", "."))
        except ValueError:
            return 0.0
    return 0.0


def dados_placa(caminho_xml: str) -> dict:
    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    ns = {"cal": "http://Petrobras/Medicao/Calibracao"}

    def get_text(elemento, tag):
        if elemento is None:
            return None
        el = elemento.find(tag, ns)
        return el.text.strip() if el is not None and el.text else None

    def extrair_bloco_simples(elemento):
        if elemento is None:
            return None

        valor = elemento.find("VALOR", ns)
        inc = elemento.find("INCERTEZA_EXP", ns)
        aprovado = elemento.find("APROVADO", ns)

        return {
            "valor": valor.text if valor is not None else None,
            "incerteza": inc.text if inc is not None else None,
            "aprovado": aprovado.text if aprovado is not None else None,
        }

    dados = {}

    dados["numero_certificado"] = get_text(root, "NUMERO_CERTIFICADO")
    dados["data_emissao"] = get_text(root, "DATA_EMISSAO")
    dados["tecnico_signatario"] = get_text(root, "TECNICO_SIGNATARIO")
    dados["tecnico_executante"] = get_text(root, "TECNICO_EXECUTANTE")

    cliente = root.find("CLIENTE", ns)
    dados["cliente"] = {
        "nome": get_text(cliente, "NOME"),
        "endereco": get_text(cliente, "ENDERECO"),
        "unidade_operacional": get_text(cliente, "UNIDADE_OPERACIONAL"),
    }


    placa = root.find("PLACA_ORIFICIO", ns)

    dados["placa"] = {
        "data_inspecao": get_text(placa, "DATA_INSPECAO"),
        "numero_serie": get_text(placa, "NUM_SERIE"),
        "tag": get_text(placa, "TAG"),
        "material": get_text(placa, "MATERIAL"),
        "coef_dilatacao": get_text(placa, "COEF_DILATACAO"),
        "norma_avaliacao": get_text(placa, "NORMA_AVALIACAO"),
        "diametro_tubulacao": get_text(placa, "DIAMETRO_TUBULACAO"),
        "beta": get_text(placa, "BETA"),
        "borda_g_sem_danos": get_text(placa, "BORDA_G_SEM_DANOS"),
        "borda_g_afiada": get_text(placa, "BORDA_G_AFIADA"),
    }

    diametro = placa.find("DIAMETRO_ORIF_COND_REF", ns)

    valor_medio = diametro.find("VALOR_MEDIO", ns)

    incerteza_elem = valor_medio.find("INCERTEZA_EXP", ns)

    dados["placa"]["diametro_orificio"] = {
        "valor_medio": {
            "valor": get_text(valor_medio, "VALOR"),
            "incerteza": get_text(valor_medio, "INCERTEZA_EXP"),
            "k": incerteza_elem.get("K") if incerteza_elem is not None else None,
            "aprovado": get_text(valor_medio, "APROVADO"),
        },
        "medidas_individuais": [
            med.text for med in diametro.findall("MEDIDA_INDIVIDUAL", ns)
        ]
    }

    esp_orificio = placa.find("ESPESSURA_ORIFICIO", ns)
    valor_medio = esp_orificio.find("VALOR_MEDIO", ns)

    dados["placa"]["espessura_orificio"] = {
        "valor_medio": {
            "valor": get_text(valor_medio, "VALOR"),
            "incerteza": get_text(valor_medio, "INCERTEZA_EXP"),
            "aprovado": get_text(valor_medio, "APROVADO"),
        },
        "medidas_individuais": [
            med.text for med in esp_orificio.findall("MEDIDA_INDIVIDUAL", ns)
        ]
    }

    esp_placa = placa.find("ESPESSURA_PLACA", ns)
    valor_medio = esp_placa.find("VALOR_MEDIO", ns)

    dados["placa"]["espessura_placa"] = {
        "valor_medio": {
            "valor": get_text(valor_medio, "VALOR"),
            "incerteza": get_text(valor_medio, "INCERTEZA_EXP"),
            "aprovado": get_text(valor_medio, "APROVADO"),
        },
        "medidas_individuais": [
            med.text for med in esp_placa.findall("MEDIDA_INDIVIDUAL", ns)
        ]
    }

    dados["placa"]["circularidade"] = extrair_bloco_simples(
        placa.find("CIRCULARIDADE_ORIF", ns)
    )

    dados["placa"]["rugosidade_montante"] = extrair_bloco_simples(
        placa.find("RUGOSIDADE_MONTANTE", ns)
    )

    dados["placa"]["rugosidade_jusante"] = extrair_bloco_simples(
        placa.find("RUGOSIDADE_JUSANTE", ns)
    )

    dados["placa"]["planeza_montante"] = extrair_bloco_simples(
        placa.find("PLANEZA_MONTANTE", ns)
    )

    dados["placa"]["angulo_chanfro"] = extrair_bloco_simples(
        placa.find("ANGULO_DE_CHANFRO", ns)
    )

    dados["placa"]["angulo_borda_g"] = extrair_bloco_simples(
        placa.find("ANGULO_BORDA_G", ns)
    )

    dados["placa"]["excentricidade"] = extrair_bloco_simples(
        placa.find("EXCENTRICIDADE", ns)
    )

    return dados

def dados_cromatografia(caminho_xml: str) -> dict:
    tree = ET.parse(caminho_xml)
    root = tree.getroot()

    def normalizar_numero(valor):
        if valor is None:
            return None
        return valor.replace(",", ".")

    def get_text(elemento, tag):
        el = elemento.find(tag)
        return el.text.strip() if el is not None and el.text else None

    dados = {}

    cabecalho = root.find("CABECALHO")
    dados["cabecalho"] = {
        "empresa": get_text(cabecalho, "EMPRESA"),
        "certificado": get_text(cabecalho, "CERTIFICADO"),
    }

    dados["componentes"] = []

    for comp in root.findall("COMPOSICAOGASES/COMPONENTE"):
        dados["componentes"].append({
            "rotulo": comp.attrib.get("rotulo"),
            "nome": get_text(comp, "NOME"),
            "molpct": normalizar_numero(get_text(comp, "MOLPCT")),
            "incerteza": normalizar_numero(get_text(comp, "INCERTEZA")),
        })

    dados["propriedades_condicao_padrao"] = []

    for prop in root.findall("PROPRIEDADESCONDICAOPADRAO/PROPRIEDADE"):
        dados["propriedades_condicao_padrao"].append({
            "nome": get_text(prop, "NOME"),
            "valor": normalizar_numero(get_text(prop, "VALOR")),
            "incerteza": normalizar_numero(get_text(prop, "INCERTEZA")),
            "referencia": prop.attrib.get("referencia"),
        })

    dados["propriedades_condicoes_amostragem"] = []

    for prop in root.findall("PROPRIEDADESCONDICOESAMOSTRAGEM/PROPRIEDADE"):
        dados["propriedades_condicoes_amostragem"].append({
            "nome": get_text(prop, "NOME"),
            "valor": normalizar_numero(get_text(prop, "VALOR")),
            "incerteza": normalizar_numero(get_text(prop, "INCERTEZA")),
            "referencia": prop.attrib.get("referencia"),
        })

    return dados


