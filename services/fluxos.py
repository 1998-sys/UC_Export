from tkinter import messagebox


def fluxo_gas(selecionar_xmls_gas, inserir_dados_operacao):

    selecionar_xmls_gas()

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