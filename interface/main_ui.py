import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import os
import sys
import traceback

from services.ci_service import identificar_tipo_ci, executar_fluxo
from writers.utils_writer import registrar_resposta
from services.fluxos import fluxo_gas, fluxo_oleo
from loaders.loader_xml import (
    dados_secundarios,
    dados_placa,
    dados_cromatografia,
    identificar_tipo_xml,
    extrair_max_pressao,
)


def resource_path(relative_path):
    """Resolve caminhos de recursos compatível com PyInstaller e modo dev."""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


ctk.set_appearance_mode("light")

ODS_RED = "#D81F3C"
ODS_RED_HOVER = "#B51A32"
ODS_BG = "#FFFFFF"
ODS_DARK = "#343A40"
ODS_TEXT = "#1A1A1A"
ODS_FRAME_LIGHT = "#F8F9FA"

FONT_FAMILY = "Segoe UI"

dados_coletados = {}


def selecionar_ci():
    """Ponto de entrada principal: abre o CI Excel, identifica o tipo (gás/óleo) e dispara o fluxo correspondente."""

    global dados_coletados
    dados_coletados.clear()

    caminho = filedialog.askopenfilename(
        title="Selecionar CI",
        filetypes=[("Excel", "*.xlsx *.xlsm *.xls")]
    )

    if not caminho:
        return

    tipo = identificar_tipo_ci(caminho)

    if tipo == "nao_suportado":
        messagebox.showwarning(
            "CI não suportado",
            "A planilha selecionada não é suportada."
        )
        return

    if tipo == "gas":
        messagebox.showinfo(
            "Tipo de CI",
            "Planilha de gás selecionada."
        )

    elif tipo == "oleo":
        messagebox.showinfo(
            "Tipo de CI",
            "Planilha de óleo selecionada."
        )

    dados_coletados["tipo_ci"] = tipo
    dados_coletados["ci"] = caminho

    iniciar_fluxo()


def selecionar_xmls_oleo():
    """
    Seleção múltipla de XMLs para o fluxo de óleo.
    Classifica cada arquivo automaticamente pelo root tag do certificado Petrobras,
    mapeando: pressao → pressao_estatica, temperatura → temperatura, termorresistencia → termoresistencia.
    """
    caminhos = filedialog.askopenfilenames(
        title="Selecionar XMLs de Calibração (Óleo)",
        filetypes=[("Arquivos XML", "*.xml")]
    )

    if not caminhos:
        return

    mapa_tipos = {
        "pressao":          "pressao_estatica",
        "temperatura":      "temperatura",
        "termorresistencia": "termoresistencia",
    }

    for caminho in caminhos:
        try:
            tipo = identificar_tipo_xml(caminho)
            chave = mapa_tipos.get(tipo)

            if chave is None:
                messagebox.showwarning(
                    "Tipo não reconhecido",
                    f"Instrumento não identificado:\n{os.path.basename(caminho)}"
                )
                continue

            dados_coletados[chave] = dados_secundarios(caminho)
            registrar_resposta(chave, True)

        except Exception as e:
            messagebox.showerror(
                "Erro ao ler XML",
                f"{os.path.basename(caminho)}\n{e}"
            )


def selecionar_xmls_gas_calibracao():
    """
    Seleção múltipla de XMLs para o fluxo de gás.
    Classifica cada arquivo pelo root tag; XMLs de pressão são ordenados pela faixa máxima
    (decrescente) e atribuídos às chaves: pressao_estatica, dpt_alta, dp_baixa, dp_media.
    Máximo de 4 XMLs de pressão — excedentes são ignorados com aviso.
    """
    caminhos = filedialog.askopenfilenames(
        title="Selecionar XMLs de Calibração (Gás)",
        filetypes=[("Arquivos XML", "*.xml")]
    )

    if not caminhos:
        return

    pressao_xmls = []  # (caminho, max_range)

    for caminho in caminhos:
        try:
            tipo = identificar_tipo_xml(caminho)

            if tipo == "placa":
                dados_coletados["placa"] = dados_placa(caminho)
                registrar_resposta("placa", True)

            elif tipo == "cromatografia":
                dados_coletados["cromatografia"] = dados_cromatografia(caminho)
                registrar_resposta("cromatografia", True)

            elif tipo == "termorresistencia":
                dados_coletados["termoresistencia"] = dados_secundarios(caminho)
                registrar_resposta("termoresistencia", True)

            elif tipo == "temperatura":
                dados_coletados["temperatura"] = dados_secundarios(caminho)
                registrar_resposta("temperatura", True)

            elif tipo == "pressao":
                # Acumula para ordenar por faixa antes de atribuir chaves
                max_range = extrair_max_pressao(caminho)
                pressao_xmls.append((caminho, max_range))

            else:
                messagebox.showwarning(
                    "Tipo não reconhecido",
                    f"Instrumento não identificado:\n{os.path.basename(caminho)}"
                )

        except Exception as e:
            messagebox.showerror(
                "Erro ao ler XML",
                f"{os.path.basename(caminho)}\n{e}"
            )

    if not pressao_xmls:
        return

    # Ordena do maior para o menor range e atribui às chaves em ordem de prioridade
    pressao_xmls.sort(key=lambda x: x[1], reverse=True)
    chaves_pressao = ["pressao_estatica", "dpt_alta", "dp_baixa", "dp_media"]

    if len(pressao_xmls) > len(chaves_pressao):
        nomes = ", ".join(os.path.basename(c) for c, _ in pressao_xmls[len(chaves_pressao):])
        messagebox.showwarning(
            "Excesso de XMLs de Pressão",
            f"Mais de 4 XMLs de pressão selecionados. Ignorando:\n{nomes}"
        )
        pressao_xmls = pressao_xmls[:len(chaves_pressao)]

    for i, (caminho, _) in enumerate(pressao_xmls):
        try:
            chave = chaves_pressao[i]
            dados_coletados[chave] = dados_secundarios(caminho)
            registrar_resposta(chave, True)
        except Exception as e:
            messagebox.showerror(
                "Erro ao ler XML",
                f"{os.path.basename(caminho)}\n{e}"
            )


def inserir_dados_operacao():
    """Janela modal para coleta de pressão e temperatura de operação — usada no fluxo de gás."""

    janela = ctk.CTkToplevel()
    janela.title("Dados de Operação")
    janela.geometry("350x400")
    janela.grab_set()

    frame = ctk.CTkFrame(janela, fg_color=ODS_FRAME_LIGHT)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    titulo = ctk.CTkLabel(
        frame,
        text="Inserir Dados de Operação",
        font=(FONT_FAMILY, 16, "bold"),
        text_color=ODS_TEXT
    )
    titulo.pack(pady=(10, 20))

    label_pressao = ctk.CTkLabel(frame, text="Pressão:", font=(FONT_FAMILY, 12))
    label_pressao.pack()

    entry_pressao = ctk.CTkEntry(frame)
    entry_pressao.pack(pady=5)

    label_temp = ctk.CTkLabel(frame, text="Temperatura:", font=(FONT_FAMILY, 12))
    label_temp.pack(pady=(10, 0))

    entry_temp = ctk.CTkEntry(frame)
    entry_temp.pack(pady=5)

    def salvar():
        """Valida os campos e persiste os dados de operação em dados_coletados."""
        pressao = entry_pressao.get()
        temperatura = entry_temp.get()

        if not pressao or not temperatura:
            messagebox.showwarning(
                "Atenção",
                "Preencha pressão e temperatura."
            )
            return

        dados_coletados["dados_operacao"] = {
            "pressao": pressao,
            "temperatura": temperatura
        }

        janela.destroy()

    botao_salvar = ctk.CTkButton(
        frame,
        text="Salvar",
        command=salvar,
        fg_color=ODS_RED,
        hover_color=ODS_RED_HOVER
    )

    botao_salvar.pack(pady=15)

    janela.wait_window()


def inserir_dados_op_oleo():
    """Janela modal para coleta dos dados de fluxo de óleo: densidade, temperatura, pressão, BSW e incerteza BSW."""

    janela = ctk.CTkToplevel()
    janela.title("Dados Operação Óleo")
    janela.geometry("350x550")
    janela.grab_set()

    frame = ctk.CTkFrame(janela, fg_color=ODS_FRAME_LIGHT)
    frame.pack(padx=20, pady=20, fill="both", expand=True)

    titulo = ctk.CTkLabel(
        frame,
        text="Inserir Dados de Fluxo de Óleo",
        font=(FONT_FAMILY, 16, "bold"),
        text_color=ODS_TEXT
    )
    titulo.pack(pady=(10, 20))

    # Densidade
    ctk.CTkLabel(frame, text="Densidade:", font=(FONT_FAMILY, 12)).pack()
    entry_densidade = ctk.CTkEntry(frame)
    entry_densidade.pack(pady=5)

    # Temperatura
    ctk.CTkLabel(frame, text="Temperatura:", font=(FONT_FAMILY, 12)).pack(pady=(10, 0))
    entry_temp = ctk.CTkEntry(frame)
    entry_temp.pack(pady=5)

    # Pressão
    ctk.CTkLabel(frame, text="Pressão:", font=(FONT_FAMILY, 12)).pack(pady=(10, 0))
    entry_pressao = ctk.CTkEntry(frame)
    entry_pressao.pack(pady=5)

    # BSW Máximo
    ctk.CTkLabel(frame, text="BSW Máximo (%):", font=(FONT_FAMILY, 12)).pack(pady=(10, 0))
    entry_bsw = ctk.CTkEntry(frame)
    entry_bsw.pack(pady=5)

    # Incerteza BSW
    ctk.CTkLabel(frame, text="Incerteza BSW (%):", font=(FONT_FAMILY, 12)).pack(pady=(10, 0))
    entry_incert_bsw = ctk.CTkEntry(frame)
    entry_incert_bsw.pack(pady=5)

    def salvar():
        """Valida todos os campos e persiste os dados de fluxo de óleo em dados_coletados."""
        densidade = entry_densidade.get()
        temperatura = entry_temp.get()
        pressao = entry_pressao.get()
        bsw = entry_bsw.get()
        inc_bsw = entry_incert_bsw.get()

        if not all([densidade, temperatura, pressao, bsw, inc_bsw]):
            messagebox.showwarning(
                "Atenção",
                "Preencha todos os campos."
            )
            return

        dados_coletados["dados_fluxo_oleo"] = {
            "densidade": densidade,
            "temperatura": temperatura,
            "pressao": pressao,
            "bsw_max": bsw,
            "incerteza_bsw": inc_bsw
        }

        janela.destroy()

    botao_salvar = ctk.CTkButton(
        frame,
        text="Salvar",
        command=salvar,
        fg_color=ODS_RED,
        hover_color=ODS_RED_HOVER
    )

    botao_salvar.pack(pady=15)

    janela.wait_window()


def iniciar_fluxo():
    """Despacha para o fluxo correto (gás/óleo) e, após coleta dos dados, chama finalizar(). Aborta se nenhum dado foi importado."""

    tipo = dados_coletados.get("tipo_ci")

    if tipo == "gas":
        fluxo_gas(selecionar_xmls_gas_calibracao, inserir_dados_operacao)

    elif tipo == "oleo":
        fluxo_oleo(selecionar_xmls_oleo, inserir_dados_op_oleo)

    else:
        messagebox.showerror("Erro", "Tipo de CI inválido.")
        return

    dados_importados = {
        k: v for k, v in dados_coletados.items()
        if k not in ("ci", "tipo_ci")
    }

    if not dados_importados:
        messagebox.showwarning(
            "Atenção",
            "Nenhum dado foi importado. Operação cancelada."
        )
        return

    finalizar()


def finalizar():
    """Envia os dados coletados para executar_fluxo(), que escreve o CI. Exibe confirmação ou erro ao usuário."""

    caminho_ci = dados_coletados.get("ci")
    tipo = dados_coletados.get("tipo_ci")
    print(dados_coletados)

    if not caminho_ci:
        messagebox.showerror("Erro", "Nenhuma planilha CI foi selecionada.")
        return

    try:

        dados_para_envio = {
            k: v for k, v in dados_coletados.items()
            if k not in ("ci", "tipo_ci")
        }

        if tipo == "gas":
            executar_fluxo(ci_path=caminho_ci,dados=dados_para_envio,tipo=tipo)

        elif tipo == "oleo":

                messagebox.showwarning(
                    "Atenção",
                    "fluxo para óleo rodando."
                )
                executar_fluxo(
                ci_path=caminho_ci,
                dados=dados_para_envio,
                tipo=tipo
            )



        else:
            messagebox.showerror("Erro", "Tipo de CI inválido.")
            return

        messagebox.showinfo(
            "Concluído",
            "Planilha atualizada com sucesso!"
        )

    except Exception as e:
        messagebox.showerror(
            "Erro",
            f"Ocorreu um erro:\n{str(e)}\n\n{traceback.format_exc()}"
        )


class App(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title('UC Export')
        self.geometry("420x550")
        self.resizable(False, False)
        self.configure(fg_color=ODS_BG)

        self._definir_logo_janela()
        self._build_ui()

    def _definir_logo_janela(self):
        """Carrega o ícone da janela via PhotoImage; falha silenciosa se o asset não existir."""

        try:

            caminho_logo = resource_path("logo/ods-logo2.png")

            img_icon = Image.open(caminho_logo)
            self.icon_img = ctk.CTkImage(img_icon, size=(32, 32))

            from tkinter import PhotoImage
            self.after(
                200,
                lambda: self.wm_iconphoto(False, PhotoImage(file=caminho_logo))
            )

        except Exception as e:
            print("Erro ao carregar ícone:", e)

    def _build_ui(self):
        """Constrói o layout principal: header com título, logo central e botão de seleção de CI."""

        header = ctk.CTkFrame(self, fg_color=ODS_RED, height=100, corner_radius=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        container = ctk.CTkFrame(header, fg_color="transparent")
        container.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            container,
            text="UC EXPORT",
            text_color="white",
            font=ctk.CTkFont(family=FONT_FAMILY, size=26, weight="bold")
        ).pack()

        ctk.CTkLabel(
            container,
            text="ODS ENERGY SOLUTIONS",
            text_color="white",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold")
        ).pack()

        self.content_frame = ctk.CTkFrame(self, fg_color=ODS_BG)
        self.content_frame.pack(fill="both", expand=True, padx=25)

        try:

            caminho_logo = resource_path("logo/logo.jpg")

            img = Image.open(caminho_logo)

            self.logo_dog = ctk.CTkImage(
                light_image=img,
                dark_image=img,
                size=(140, 140)
            )

            self.dog_label = ctk.CTkLabel(
                self.content_frame,
                image=self.logo_dog,
                text=""
            )

            self.dog_label.pack(pady=(25, 15))

        except Exception as e:
            print("Erro ao carregar logo:", e)

        self.btn_ci = ctk.CTkButton(
            self.content_frame,
            text="SELECIONAR CI",
            width=240,
            height=60,
            fg_color=ODS_RED,
            hover_color=ODS_RED_HOVER,
            font=(FONT_FAMILY, 14, "bold"),
            corner_radius=12,
            command=selecionar_ci
        )

        self.btn_ci.pack(pady=15)

        footer = ctk.CTkFrame(self, fg_color=ODS_DARK, height=35, corner_radius=0)
        footer.pack(fill="x", side="bottom")

        ctk.CTkLabel(
            footer,
            text="Developed by M. Bandeira, G. Machado © 2026",
            text_color="#BDC3C7",
            font=(FONT_FAMILY, 10)
        ).pack(expand=True)


def run():
    """Instancia e executa o loop principal da aplicação."""
    app = App()
    app.mainloop()
