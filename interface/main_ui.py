import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
import os
import sys

from services.ci_service import identificar_tipo_ci, executar_fluxo
from services.validacoes import validar_ordem_dpts
from writers.utils_writer import registrar_resposta
from services.fluxos import fluxo_gas, fluxo_oleo
from loaders.loader_xml import (
    dados_secundarios,
    dados_placa,
    dados_cromatografia
)


def resource_path(relative_path):
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


def perguntar_xml(pergunta, chave, tipo_loader="certificado"):
  
    resposta = messagebox.askyesno(
        "Inserir Dados",
        f"Deseja inserir dados de {pergunta}?"
    )

    if not resposta:
        registrar_resposta(chave, False)
        return
    
    caminho_xml = filedialog.askopenfilename(
        title=f"Selecionar XML - {pergunta}",
        filetypes=[("Arquivos XML", "*.xml")]
    )


    if not caminho_xml:
        registrar_resposta(chave, False)
        return

    try:

        if tipo_loader == "certificado":
            dados = dados_secundarios(caminho_xml)

        elif tipo_loader == "placa":
            dados = dados_placa(caminho_xml)

        elif tipo_loader == "cromatografia":
            dados = dados_cromatografia(caminho_xml)

        dados_coletados[chave] = dados
        registrar_resposta(chave, True)

    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler XML:\n{e}")


def inserir_dados_operacao():

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


def iniciar_fluxo():

    tipo = dados_coletados.get("tipo_ci")
    
    if tipo == "gas":
        fluxo_gas(perguntar_xml, inserir_dados_operacao)

    elif tipo == "oleo":
        fluxo_oleo(perguntar_xml, inserir_dados_operacao)

    else:
        messagebox.showerror("Erro", "Tipo de CI inválido.")
        return

    finalizar()


def finalizar():

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

            valido, mensagem = validar_ordem_dpts(dados_para_envio)

            if not valido:
                messagebox.showwarning(
                    "Atenção - Ordem dos DPTs",
                    mensagem
                )
                return
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
            f"Ocorreu um erro:\n{str(e)}"
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

    app = App()
    app.mainloop()