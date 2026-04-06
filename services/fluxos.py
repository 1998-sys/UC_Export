from tkinter import messagebox


def fluxo_gas(perguntar_xml, inserir_dados_operacao):

    perguntar_xml("Placa de Orifício", "placa", "placa")
    perguntar_xml("Cromatografia", "cromatografia", "cromatografia")
    perguntar_xml("Transmissor de Pressão Diferencial Alta", "dpt_alta")
    perguntar_xml("Transmissor de Pressão Diferencial Média", "dp_media")
    perguntar_xml("Transmissor de Pressão Diferencial Baixa", "dp_baixa")
    perguntar_xml("Pressão Estática", "pressao_estatica")
    perguntar_xml("Transmissor de Temperatura", "temperatura")
    perguntar_xml("Termorresistência", "termoresistencia")

    perguntar_dados_operacao(inserir_dados_operacao)


def fluxo_oleo(selecionar_xmls_oleo, inserir_dados_operacao_oleo):

    selecionar_xmls_oleo()

    perguntar_dados_operacao(inserir_dados_operacao_oleo)


def perguntar_dados_operacao(inserir_dados_operacao):

    resposta = messagebox.askyesno(
        "Dados de Operação",
        "Deseja inserir dados de operação?"
    )

    if resposta:
        inserir_dados_operacao()