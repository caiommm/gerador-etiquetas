import os
import sys

# Desativa aceleração de hardware/D3D para evitar erros de DLL de gráficos em Win 7
os.environ["TK_SILENT_INIT"] = "1"

import json
import re
import tempfile
import unicodedata
import pandas as pd
import docx
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

import barcode
from barcode.writer import ImageWriter
import qrcode
from PIL import Image

import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk

# Tratamento seguro para winsound no Win 7
try:
    import winsound
except Exception:
    winsound = None

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

CONFIG_FILE = "config_etiquetas.json"

MODELOS_PIMACO = {
    "Pimaco A4251 (65 etq - 3.82 x 2.12 cm)": {
        "cols": 5,
        "width_cm": 3.82,
        "height_cm": 2.12,
        "top_m": 1.07,
        "bot_m": 1.07,
        "left_m": 0.4,
        "right_m": 0.4,
        "row_height": 1200,
        "default_f_top": 5.5,
        "default_f_desc": 4.5,
        "default_entrelinha": "0.85",
        "default_f_num_pct": "30",
        "default_barcode_width_cm": "2.6",
        "default_barcode_height_mm": "1.8",
        "default_text_distance": "0.1",
        "default_space_before_bc": "0",
        "default_space_after_top": "0.2",
        "default_pad_top": "0.0",
        "default_pad_left": "0.0",
    },
    "Pimaco 6080 / 6280 (21 etq - 6.35 x 3.81 cm)": {
        "cols": 3,
        "width_cm": 6.35,
        "height_cm": 3.81,
        "top_m": 1.5,
        "bot_m": 1.5,
        "left_m": 0.7,
        "right_m": 0.7,
        "row_height": 2160,
        "default_f_top": 8.0,
        "default_f_desc": 7.0,
        "default_entrelinha": "0.90",
        "default_f_num_pct": "35",
        "default_barcode_width_cm": "3.8",
        "default_barcode_height_mm": "3.5",
        "default_text_distance": "0.3",
        "default_space_before_bc": "1",
        "default_space_after_top": "1.0",
        "default_pad_top": "0.0",
        "default_pad_left": "0.0",
    },
}


def remover_acentos(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize("NFKD", str(texto))
    texto_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return (
        texto_sem_acento.replace("Ç", "C")
        .replace("ç", "c")
        .encode("ascii", "ignore")
        .decode("utf-8")
    )


def aplicar_recuo_esquerda(paragraph, pad_left_pt):
    """Aplica recuo à esquerda via XML puro (permite valores negativos/DXA)."""
    val_dxa = int(pad_left_pt * 20)
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith('ind'):
            pPr.remove(child)
    ind = parse_xml(f'<w:ind {nsdecls("w")} w:left="{val_dxa}"/>')
    pPr.append(ind)


def aplicar_espacamento_paragrafo(paragraph, space_before_pt=0, space_after_pt=0, line_spacing=0.85):
    """Aplica espaçamento antes/depois e entrelinha via XML, tratando valores negativos para não quebrar o Word."""
    pPr = paragraph._p.get_or_add_pPr()
    for child in list(pPr):
        if child.tag.endswith('spacing'):
            pPr.remove(child)

    before_dxa = max(0, int(space_before_pt * 20))
    after_dxa = max(0, int(space_after_pt * 20))
    line_val = int(line_spacing * 240)

    spacing = parse_xml(
        f'<w:spacing {nsdecls("w")} '
        f'w:before="{before_dxa}" '
        f'w:after="{after_dxa}" '
        f'w:line="{line_val}" '
        f'w:lineRule="auto"/>'
    )
    pPr.append(spacing)


def gerar_imagem_barcode(codigo_str, font_size_pt=6, module_height_mm=2.0, text_distance=0.2):
    codigo_limpo = re.sub(r'[^A-Za-z0-9]', '', str(codigo_str).strip())
    if not codigo_limpo:
        codigo_limpo = "7896067228753"

    # Tenta localizar uma fonte padrão do Windows para evitar erro de DLL/Font no PyInstaller
    caminho_fonte = os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "arial.ttf")
    if not os.path.exists(caminho_fonte):
        caminho_fonte = None

    writer_options = {
        "module_width": 0.2,
        "module_height": max(1.5, float(module_height_mm)),
        "quiet_zone": 1.0,
        "font_size": max(4, int(font_size_pt)),
        "text_distance": max(0.1, float(text_distance)),
        "write_text": True,
    }
    
    if caminho_fonte:
        writer_options["font_path"] = caminho_fonte

    try:
        if codigo_limpo.isdigit() and len(codigo_limpo) in [12, 13]:
            if len(codigo_limpo) == 12:
                codigo_limpo = codigo_limpo.zfill(13)
            BARCODE_CLASS = barcode.get_barcode_class("ean13")
        else:
            BARCODE_CLASS = barcode.get_barcode_class("code128")

        bc = BARCODE_CLASS(codigo_limpo, writer=ImageWriter())
        pasta_temp = tempfile.gettempdir()
        caminho_base = os.path.join(pasta_temp, f"temp_bar_{codigo_limpo}")
        arquivo_salvo = bc.save(caminho_base, options=writer_options)

        if not arquivo_salvo.endswith('.png'):
            arquivo_salvo += '.png'

        return arquivo_salvo
    except Exception as e:
        print(f"Erro ao gerar código de barras: {e}")
        # Fallback de segurança em caso de erro na fonte
        try:
            CODE128 = barcode.get_barcode_class("code128")
            bc = CODE128(codigo_limpo, writer=ImageWriter())
            pasta_temp = tempfile.gettempdir()
            caminho_base = os.path.join(pasta_temp, f"temp_bar_{codigo_limpo}")
            writer_options["write_text"] = False  # Desativa texto se der erro de fonte
            return bc.save(caminho_base, options=writer_options)
        except Exception:
            return None


def gerar_imagem_qrcode(codigo_str):
    codigo_limpo = str(codigo_str).strip()
    if not codigo_limpo:
        codigo_limpo = "7896067228753"

    try:
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_M,
            box_size=8,
            border=1,
        )
        qr.add_data(codigo_limpo)
        qr.make(fit=True)

        img = qr.make_image(fill_color="black", back_color="white")
        pasta_temp = tempfile.gettempdir()
        codigo_nome = re.sub(r'[^A-Za-z0-9]', '', codigo_limpo) or "qrcode"
        caminho_salvo = os.path.join(pasta_temp, f"temp_qr_{codigo_nome}.png")
        img.save(caminho_salvo)

        return caminho_salvo
    except Exception as e:
        print(f"Erro ao gerar QR Code: {e}")
        return None


def gerar_imagem_codigo(tipo, codigo_str, font_size_pt=6, module_height_mm=2.0, text_distance=0.2):
    if tipo == "QR Code":
        return gerar_imagem_qrcode(codigo_str)
    else:
        return gerar_imagem_barcode(codigo_str, font_size_pt, module_height_mm, text_distance)


def aplicar_bordas_tabela(tabela, exibir_bordas=False):
    tblPr = tabela._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)

    val_borda = "single" if exibir_bordas else "none"
    cor_borda = "CCCCCC" if exibir_bordas else "auto"
    sz_borda = "4" if exibir_bordas else "0"

    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        f'  <w:left w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        f'  <w:bottom w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        f'  <w:right w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        f'  <w:insideH w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        f'  <w:insideV w:val="{val_borda}" w:sz="{sz_borda}" w:space="0" w:color="{cor_borda}"/>\n'
        '</w:tblBorders>'
    )
    tblPr.append(borders)


class AppEtiquetas(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Gerador de Etiquetas FL 38 Niteroi")
        self.geometry("1300x900")
        self.minsize(1100, 750)

        self.caminho_planilha = ""
        self.df_carregado = None
        self._job_preview = None
        self.barcode_preview_img = None

        self.title_label = ctk.CTkLabel(
            self, text="Gerador de Etiquetas FL 38 Niteroi ;)", font=("Arial", 20, "bold")
        )
        self.title_label.pack(pady=(10, 5), side="top")

        self.bottom_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.bottom_frame.pack(side="bottom", fill="x", padx=15, pady=(5, 10))

        self.lbl_status = ctk.CTkLabel(self.bottom_frame, text="Aguardando planilha...", text_color="gray")
        self.lbl_status.pack(pady=(2, 2))

        self.progress_bar = ctk.CTkProgressBar(self.bottom_frame, width=700)
        self.progress_bar.pack(pady=4)
        self.progress_bar.set(0)

        self.btn_gerar = ctk.CTkButton(
            self.bottom_frame,
            text="📄 Gerar Documento Definitivo no Word (.docx)",
            command=self.gerar_etiquetas_word,
            fg_color="green",
            hover_color="darkgreen",
            font=("Arial", 14, "bold"),
            height=42,
        )
        self.btn_gerar.pack(pady=(6, 2))

        self.main_container = ctk.CTkFrame(self)
        self.main_container.pack(fill="both", expand=True, padx=15, pady=5)

        self.main_container.grid_columnconfigure(0, weight=6)
        self.main_container.grid_columnconfigure(1, weight=4, minsize=420)
        self.main_container.grid_rowconfigure(0, weight=1)

        self.left_frame = ctk.CTkScrollableFrame(self.main_container, label_text="Configurações e Dados")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)

        self.right_frame = ctk.CTkFrame(self.main_container)
        self.right_frame.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

        self.montar_painel_esquerdo()
        self.montar_painel_direito()

        self.carregar_perfil_config()
        self.atualizar_preview_card()

    def criar_campo_map(self, parent, label_text, row, col):
        lbl = ctk.CTkLabel(parent, text=label_text, font=("Arial", 11))
        lbl.grid(row=row, column=col, padx=4, pady=3, sticky="e")

        combo = ctk.CTkOptionMenu(
            parent, values=["Nenhum"], width=130, command=lambda _: self.agendar_atualizacao_preview()
        )
        combo.grid(row=row, column=col + 1, padx=4, pady=3, sticky="w")
        return combo

    def criar_spinbox(self, parent, label_text, row, col, valor_inicial="0.0", passo=0.5, permitir_negativo=True):
        lbl = ctk.CTkLabel(parent, text=label_text, font=("Arial", 11))
        lbl.grid(row=row, column=col, padx=4, pady=2, sticky="e")

        sub_frame = ctk.CTkFrame(parent, fg_color="transparent")
        sub_frame.grid(row=row, column=col + 1, padx=4, pady=2, sticky="w")

        btn_menos = ctk.CTkButton(
            sub_frame,
            text="-",
            width=24,
            height=26,
            font=("Arial", 12, "bold"),
            command=lambda: self.alterar_valor_spinbox(entry, -passo, permitir_negativo),
        )
        btn_menos.pack(side="left", padx=(0, 2))

        entry = ctk.CTkEntry(sub_frame, width=58, height=26, justify="center")
        entry.insert(0, str(valor_inicial))
        entry.pack(side="left")
        entry.bind("<KeyRelease>", self.agendar_atualizacao_preview)

        btn_mais = ctk.CTkButton(
            sub_frame,
            text="+",
            width=24,
            height=26,
            font=("Arial", 12, "bold"),
            command=lambda: self.alterar_valor_spinbox(entry, passo, permitir_negativo),
        )
        btn_mais.pack(side="left", padx=(2, 0))

        return entry

    def alterar_valor_spinbox(self, entry, delta, permitir_negativo=True):
        try:
            val_atual = float(entry.get().replace(",", ".").strip())
        except ValueError:
            val_atual = 0.0

        novo_val = val_atual + delta

        if not permitir_negativo and novo_val < 0:
            novo_val = 0.0

        if novo_val.is_integer():
            str_val = str(int(novo_val))
        else:
            str_val = f"{novo_val:.2f}".rstrip("0").rstrip(".")

        entry.delete(0, "end")
        entry.insert(0, str_val)
        self.agendar_atualizacao_preview()

    def agendar_atualizacao_preview(self, *args):
        if self._job_preview:
            self.after_cancel(self._job_preview)
        self._job_preview = self.after(150, self.atualizar_preview_card)

    def obter_valor_float(self, entry_widget, padrao=0.0):
        try:
            val = str(entry_widget.get()).replace(",", ".").strip()
            return float(val) if val else padrao
        except Exception:
            return padrao

    def montar_painel_esquerdo(self):
        frame_top = ctk.CTkFrame(self.left_frame)
        frame_top.pack(fill="x", padx=10, pady=5)

        self.btn_planilha = ctk.CTkButton(
            frame_top, text="1. Selecionar Planilha (.xlsx / .csv)", command=self.selecionar_planilha
        )
        self.btn_planilha.pack(side="left", padx=10, pady=8)

        self.lbl_planilha = ctk.CTkLabel(
            frame_top, text="Nenhum arquivo carregado", text_color="gray"
        )
        self.lbl_planilha.pack(side="left", padx=10)

        frame_modelo = ctk.CTkFrame(self.left_frame)
        frame_modelo.pack(fill="x", padx=10, pady=5)

        ctk.CTkLabel(frame_modelo, text="Modelo Pimaco:", font=("Arial", 11, "bold")).pack(side="left", padx=8, pady=5)
        self.combo_modelo_pimaco = ctk.CTkOptionMenu(
            frame_modelo,
            values=list(MODELOS_PIMACO.keys()),
            command=lambda _: self.ao_mudar_modelo(),
            width=280,
        )
        self.combo_modelo_pimaco.pack(side="left", padx=5, pady=5)

        self.chk_incluir_bordas = ctk.CTkCheckBox(
            frame_modelo, text="Mostrar Bordas", command=self.agendar_atualizacao_preview
        )
        self.chk_incluir_bordas.pack(side="left", padx=15, pady=5)

        frame_map = ctk.CTkFrame(self.left_frame)
        frame_map.pack(fill="x", padx=10, pady=5)

        lbl_map = ctk.CTkLabel(
            frame_map, text="2. Mapeamento de Colunas e Ajustes Finos:", font=("Arial", 12, "bold")
        )
        lbl_map.grid(row=0, column=0, columnspan=4, sticky="w", padx=8, pady=4)

        self.combo_codigo = self.criar_campo_map(frame_map, "1º Código:", 1, 0)
        self.combo_rua = self.criar_campo_map(frame_map, "2º Rua/Local:", 1, 2)
        self.combo_linha = self.criar_campo_map(frame_map, "3º Linha/Setor:", 2, 0)
        self.combo_desc = self.criar_campo_map(frame_map, "4º Descrição:", 2, 2)
        self.combo_ean = self.criar_campo_map(frame_map, "5º EAN (Barras):", 3, 0)
        self.combo_qtde = self.criar_campo_map(frame_map, "Coluna Qtde:", 3, 2)

        lbl_tipo_codigo = ctk.CTkLabel(frame_map, text="Tipo de Código:", font=("Arial", 11, "bold"))
        lbl_tipo_codigo.grid(row=4, column=0, padx=4, pady=3, sticky="e")

        self.combo_tipo_codigo = ctk.CTkOptionMenu(
            frame_map,
            values=["Código de Barras (Zebra)", "QR Code"],
            width=140,
            command=lambda _: self.agendar_atualizacao_preview()
        )
        self.combo_tipo_codigo.grid(row=4, column=1, padx=4, pady=3, sticky="w")

        self.entry_qtde_manual = self.criar_spinbox(
            frame_map, "OU Qtde Fixa/Manual:", 4, 2, valor_inicial="1", passo=1, permitir_negativo=False
        )

        self.entry_posicao_inicial = self.criar_spinbox(
            frame_map, "Etiqueta Inicial (1-65):", 5, 0, valor_inicial="1", passo=1, permitir_negativo=False
        )

        lbl_sub = ctk.CTkLabel(
            frame_map, text="Controles de Fonte, Tamanho e Alinhamento Físico (+ / -):", font=("Arial", 11, "bold"), text_color="#1f538d"
        )
        lbl_sub.grid(row=6, column=0, columnspan=4, sticky="w", padx=8, pady=(8, 4))

        self.entry_fonte_top = self.criar_spinbox(
            frame_map, "Fonte Topo (pt):", 7, 0, valor_inicial="5.5", passo=0.5, permitir_negativo=False
        )

        self.entry_fonte_desc = self.criar_spinbox(
            frame_map, "Fonte Desc. (pt):", 7, 2, valor_inicial="4.5", passo=0.5, permitir_negativo=False
        )

        self.entry_entrelinha = self.criar_spinbox(
            frame_map, "Entrelinha Texto:", 8, 0, valor_inicial="0.85", passo=0.05, permitir_negativo=False
        )

        self.entry_espaco_topo_desc = self.criar_spinbox(
            frame_map, "Ajuste Topo/Desc (pt):", 8, 2, valor_inicial="0.2", passo=0.5, permitir_negativo=True
        )

        self.entry_fonte_numeral = self.criar_spinbox(
            frame_map, "Fonte Numeral (%):", 9, 0, valor_inicial="30", passo=5, permitir_negativo=False
        )

        self.entry_largura_barras = self.criar_spinbox(
            frame_map, "Largura Barras/QR (cm):", 9, 2, valor_inicial="2.6", passo=0.1, permitir_negativo=False
        )

        self.entry_altura_barras = self.criar_spinbox(
            frame_map, "Altura Barras (mm):", 10, 0, valor_inicial="1.8", passo=0.5, permitir_negativo=False
        )

        self.entry_distancia_numeral = self.criar_spinbox(
            frame_map, "Dist. Núm/Barra (pt):", 10, 2, valor_inicial="0.1", passo=0.1, permitir_negativo=True
        )

        self.entry_espaco_antes_bc = self.criar_spinbox(
            frame_map, "Espaço Antes Código (pt):", 11, 0, valor_inicial="0", passo=1, permitir_negativo=True
        )

        self.entry_pad_top = self.criar_spinbox(
            frame_map, "Margem Topo (pt):", 11, 2, valor_inicial="0.0", passo=0.5, permitir_negativo=True
        )

        self.entry_pad_left = self.criar_spinbox(
            frame_map, "Margem Esquerda (pt):", 12, 0, valor_inicial="0.0", passo=0.5, permitir_negativo=True
        )

        btn_salvar_perfil = ctk.CTkButton(
            frame_map, text="💾 Salvar Perfil", command=self.salvar_perfil_config, width=120
        )
        btn_salvar_perfil.grid(row=12, column=2, columnspan=2, padx=4, pady=(6, 8))

        frame_tree = ctk.CTkFrame(self.left_frame)
        frame_tree.pack(fill="x", expand=False, padx=10, pady=5)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=18)
        self.style.configure("Treeview.Heading", background="#1f1f1f", foreground="white", font=("Arial", 8, "bold"))

        scroll_x = ttk.Scrollbar(frame_tree, orient="horizontal")
        scroll_y = ttk.Scrollbar(frame_tree, orient="vertical")

        self.tree = ttk.Treeview(
            frame_tree, height=5, show="headings", xscrollcommand=scroll_x.set, yscrollcommand=scroll_y.set
        )
        scroll_x.config(command=self.tree.xview)
        scroll_y.config(command=self.tree.yview)

        scroll_x.pack(side="bottom", fill="x")
        scroll_y.pack(side="right", fill="y")
        self.tree.pack(fill="x", expand=True)

        self.tree.bind("<<TreeviewSelect>>", lambda e: self.agendar_atualizacao_preview())

    def ao_mudar_modelo(self):
        cfg = MODELOS_PIMACO.get(self.combo_modelo_pimaco.get())
        if cfg:
            def atualizar_entry(entry, val):
                entry.delete(0, "end")
                entry.insert(0, str(val))

            atualizar_entry(self.entry_fonte_top, cfg["default_f_top"])
            atualizar_entry(self.entry_fonte_desc, cfg["default_f_desc"])
            atualizar_entry(self.entry_entrelinha, cfg.get("default_entrelinha", "0.85"))
            atualizar_entry(self.entry_espaco_topo_desc, cfg.get("default_space_after_top", "0.2"))
            atualizar_entry(self.entry_fonte_numeral, cfg.get("default_f_num_pct", "30"))
            atualizar_entry(self.entry_largura_barras, cfg.get("default_barcode_width_cm", "2.6"))
            atualizar_entry(self.entry_altura_barras, cfg.get("default_barcode_height_mm", "1.8"))
            atualizar_entry(self.entry_distancia_numeral, cfg.get("default_text_distance", "0.1"))
            atualizar_entry(self.entry_espaco_antes_bc, cfg.get("default_space_before_bc", "0"))
            atualizar_entry(self.entry_pad_top, cfg.get("default_pad_top", "0.0"))
            atualizar_entry(self.entry_pad_left, cfg.get("default_pad_left", "0.0"))
        self.atualizar_preview_card()

    def montar_painel_direito(self):
        lbl_preview_title = ctk.CTkLabel(
            self.right_frame, text="⚡ Pré-Visualização Dinâmica em Tempo Real", font=("Arial", 13, "bold")
        )
        lbl_preview_title.pack(pady=(12, 5))

        self.container_card = ctk.CTkFrame(self.right_frame, fg_color="#1e1e1e")
        self.container_card.pack(fill="both", expand=True, padx=10, pady=10)

        self.card_etiqueta = ctk.CTkFrame(
            self.container_card, width=320, height=180, fg_color="white", corner_radius=2, border_width=1, border_color="#333333"
        )
        self.card_etiqueta.pack(expand=True, anchor="center")
        self.card_etiqueta.pack_propagate(False)

        self.inner_card = ctk.CTkFrame(self.card_etiqueta, fg_color="transparent")
        self.inner_card.pack(fill="both", expand=True)

        self.lbl_card_topo = ctk.CTkLabel(
            self.inner_card, text="1JX51015VVRDB (M) (CL)", font=("Arial", 8, "bold"), text_color="black"
        )
        self.lbl_card_topo.pack(pady=(4, 1))

        self.lbl_card_desc = ctk.CTkLabel(
            self.inner_card, text="SMARTPH 15 PRO 5G 512GB VRD", font=("Arial", 7), text_color="black", wraplength=280, justify="center"
        )
        self.lbl_card_desc.pack(pady=1)

        self.lbl_card_barcode_img = ctk.CTkLabel(self.inner_card, text="")
        self.lbl_card_barcode_img.pack(pady=(0, 2))

    def salvar_perfil_config(self):
        dados = {
            "fonte_topo": self.entry_fonte_top.get(),
            "fonte_desc": self.entry_fonte_desc.get(),
            "entrelinha": self.entry_entrelinha.get(),
            "espaco_topo_desc": self.entry_espaco_topo_desc.get(),
            "fonte_numeral": self.entry_fonte_numeral.get(),
            "largura_barras_cm": self.entry_largura_barras.get(),
            "altura_barras_mm": self.entry_altura_barras.get(),
            "distancia_numeral": self.entry_distancia_numeral.get(),
            "espaco_antes_bc": self.entry_espaco_antes_bc.get(),
            "pad_top": self.entry_pad_top.get(),
            "pad_left": self.entry_pad_left.get(),
            "modelo_pimaco": self.combo_modelo_pimaco.get(),
            "qtde_manual": self.entry_qtde_manual.get(),
            "posicao_inicial": self.entry_posicao_inicial.get(),
            "tipo_codigo": self.combo_tipo_codigo.get(),
            "incluir_bordas": self.chk_incluir_bordas.get(),
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(dados, f, indent=4)
            messagebox.showinfo("Perfil Salvo", "Configurações salvas com sucesso!", parent=self)
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o perfil: {e}", parent=self)

    def carregar_perfil_config(self):
        if not os.path.exists(CONFIG_FILE):
            return
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                dados = json.load(f)

            def carregar_val(entry, chave, padrao):
                entry.delete(0, "end")
                entry.insert(0, str(dados.get(chave, padrao)))

            carregar_val(self.entry_fonte_top, "fonte_topo", "5.5")
            carregar_val(self.entry_fonte_desc, "fonte_desc", "4.5")
            carregar_val(self.entry_entrelinha, "entrelinha", "0.85")
            carregar_val(self.entry_espaco_topo_desc, "espaco_topo_desc", "0.2")
            carregar_val(self.entry_fonte_numeral, "fonte_numeral", "30")
            carregar_val(self.entry_largura_barras, "largura_barras_cm", "2.6")
            carregar_val(self.entry_altura_barras, "altura_barras_mm", "1.8")
            carregar_val(self.entry_distancia_numeral, "distancia_numeral", "0.1")
            carregar_val(self.entry_espaco_antes_bc, "espaco_antes_bc", "0")
            carregar_val(self.entry_pad_top, "pad_top", "0.0")
            carregar_val(self.entry_pad_left, "pad_left", "0.0")

            if "modelo_pimaco" in dados and dados["modelo_pimaco"] in MODELOS_PIMACO:
                self.combo_modelo_pimaco.set(dados["modelo_pimaco"])
            if "qtde_manual" in dados:
                carregar_val(self.entry_qtde_manual, "qtde_manual", "1")
            if "posicao_inicial" in dados:
                carregar_val(self.entry_posicao_inicial, "posicao_inicial", "1")
            if "tipo_codigo" in dados:
                self.combo_tipo_codigo.set(dados["tipo_codigo"])
            if "incluir_bordas" in dados:
                if dados["incluir_bordas"]:
                    self.chk_incluir_bordas.select()
                else:
                    self.chk_incluir_bordas.deselect()

        except Exception:
            pass

    def selecionar_planilha(self):
        caminho = filedialog.askopenfilename(filetypes=[("Planilhas Excel/CSV", "*.xlsx *.csv")], parent=self)
        if caminho:
            self.caminho_planilha = caminho
            self.lbl_planilha.configure(text=os.path.basename(caminho), text_color="white")

            try:
                if caminho.endswith(".csv"):
                    self.df_carregado = pd.read_csv(caminho)
                else:
                    xls = pd.ExcelFile(caminho)
                    if "Etiquetas" in xls.sheet_names:
                        self.df_carregado = pd.read_excel(caminho, sheet_name="Etiquetas")
                    else:
                        self.df_carregado = pd.read_excel(caminho, sheet_name=0)

                colunas_originais = list(self.df_carregado.columns)
                cols = ["Nenhum"] + colunas_originais

                for combo in [self.combo_codigo, self.combo_rua, self.combo_linha, self.combo_desc, self.combo_ean, self.combo_qtde]:
                    combo.configure(values=cols)
                    combo.set("Nenhum")

                for col in colunas_originais:
                    c_limpo = str(col).strip().upper()
                    if c_limpo in ["CODIGO", "CÓDIGO", "SKU", "PRODUTO"]:
                        self.combo_codigo.set(col)
                    elif c_limpo in ["RUA", "ENDERECO", "ENDEREÇO", "LOCAL"]:
                        self.combo_rua.set(col)
                    elif c_limpo.startswith("LINHA") or c_limpo in ["SETOR", "OBS"]:
                        self.combo_linha.set(col)
                    elif c_limpo in ["DESCRICAO", "DESCRIÇÃO", "DESC", "NOME"]:
                        self.combo_desc.set(col)
                    elif c_limpo in ["EAN", "BARRA", "BARRAS", "GTIN"]:
                        self.combo_ean.set(col)
                    elif c_limpo in ["QTDE", "QTD", "QUANTIDADE"]:
                        self.combo_qtde.set(col)

                self.tree["columns"] = colunas_originais
                self.tree.delete(*self.tree.get_children())

                for col in colunas_originais:
                    self.tree.heading(col, text=col)
                    self.tree.column(col, width=110, anchor="center")

                for _, row in self.df_carregado.head(10).iterrows():
                    self.tree.insert("", "end", values=list(row))

                if len(self.tree.get_children()) > 0:
                    self.tree.selection_set(self.tree.get_children()[0])

                self.lbl_status.configure(text=f"Planilha carregada ({len(self.df_carregado)} registros).", text_color="lightgreen")
                self.atualizar_preview_card()

            except Exception as e:
                messagebox.showerror("Erro ao carregar", f"Falha ao ler planilha: {e}", parent=self)

    def atualizar_preview_card(self):
        val_codigo = "1JX51015VVRDB"
        val_rua = "M"
        val_linha = "CL"
        val_desc = "SMARTPH 15 PRO 5G 512GB VRD"
        val_ean = "6941764472341"

        if self.df_carregado is not None and not self.df_carregado.empty:
            selected = self.tree.selection()
            if selected:
                row_vals = self.tree.item(selected[0])["values"]
                cols = list(self.df_carregado.columns)
                row_dict = dict(zip(cols, row_vals))
            else:
                row_dict = self.df_carregado.iloc[0].to_dict()

            def pegar_val(combo):
                c = combo.get()
                if c != "Nenhum" and c in row_dict and pd.notna(row_dict[c]):
                    v = str(row_dict[c]).strip()
                    return v[:-2] if v.endswith(".0") else v
                return ""

            v_cod = pegar_val(self.combo_codigo)
            v_rua = pegar_val(self.combo_rua)
            v_lin = pegar_val(self.combo_linha)
            v_dsc = pegar_val(self.combo_desc)
            v_ean = pegar_val(self.combo_ean)

            if v_cod or v_rua or v_lin:
                val_codigo = v_cod
                val_rua = v_rua
                val_linha = v_lin
            if v_dsc:
                val_desc = v_dsc
            if v_ean:
                val_ean = v_ean

        f_top = self.obter_valor_float(self.entry_fonte_top, padrao=5.5)
        f_desc = self.obter_valor_float(self.entry_fonte_desc, padrao=4.5)
        espaco_topo_desc_val = self.obter_valor_float(self.entry_espaco_topo_desc, padrao=0.2)
        pct_num = self.obter_valor_float(self.entry_fonte_numeral, padrao=30.0)
        largura_barras_cm = self.obter_valor_float(self.entry_largura_barras, padrao=2.6)
        altura_barras_mm = self.obter_valor_float(self.entry_altura_barras, padrao=1.8)
        distancia_numeral_val = self.obter_valor_float(self.entry_distancia_numeral, padrao=0.1)
        espaco_antes_bc = self.obter_valor_float(self.entry_espaco_antes_bc, padrao=0.0)
        pad_top_val = int(self.obter_valor_float(self.entry_pad_top, padrao=0.0))
        pad_left_val = int(self.obter_valor_float(self.entry_pad_left, padrao=0.0))

        f_num_pt = max(4.0, (pct_num / 100.0) * 20.0)

        cfg_modelo = MODELOS_PIMACO.get(self.combo_modelo_pimaco.get(), MODELOS_PIMACO["Pimaco A4251 (65 etq - 3.82 x 2.12 cm)"])
        card_w = int(cfg_modelo["width_cm"] * 80)
        card_h = int(cfg_modelo["height_cm"] * 85)
        self.card_etiqueta.configure(width=max(260, card_w), height=max(150, card_h))

        b_width = 1 if self.chk_incluir_bordas.get() == 1 else 0
        self.card_etiqueta.configure(border_width=b_width)

        pad_left_pos = max(0, pad_left_val)
        pad_left_neg = abs(pad_left_val) if pad_left_val < 0 else 0

        pad_top_pos = max(0, pad_top_val)
        pad_top_neg = abs(pad_top_val) if pad_top_val < 0 else 0

        self.inner_card.pack_configure(
            padx=(pad_left_pos, pad_left_neg),
            pady=(pad_top_pos, pad_top_neg)
        )

        topo_p = []
        if val_codigo: topo_p.append(val_codigo)
        if val_rua: topo_p.append(f"({val_rua})")
        if val_linha: topo_p.append(f"({val_linha})")
        str_topo = " ".join(topo_p) if topo_p else "1JX51015VVRDB (M) (CL)"

        self.lbl_card_topo.configure(text=str_topo, font=("Arial", max(5, int(f_top * 1.2)), "bold"))
        self.lbl_card_topo.pack_configure(pady=(4, max(0, int(espaco_topo_desc_val))))

        self.lbl_card_desc.configure(text=val_desc, font=("Arial", max(4, int(f_desc * 1.1))))

        tipo_code = self.combo_tipo_codigo.get()
        val_ean_limpo = remover_acentos(val_ean).strip() if val_ean else "6941764472341"
        caminho_bc = gerar_imagem_codigo(
            tipo=tipo_code,
            codigo_str=val_ean_limpo,
            font_size_pt=f_num_pt,
            module_height_mm=altura_barras_mm,
            text_distance=distancia_numeral_val,
        )

        if caminho_bc and os.path.exists(caminho_bc):
            pil_img = Image.open(caminho_bc)
            target_w = int(largura_barras_cm * 80)
            aspect_ratio = pil_img.height / pil_img.width
            target_h = int(target_w * aspect_ratio)

            ctk_bc_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_w, max(18, target_h)))
            self.lbl_card_barcode_img.configure(image=ctk_bc_img, text="")

            pad_top = int(espaco_antes_bc)
            if pad_top >= 0:
                self.lbl_card_barcode_img.pack_configure(pady=(pad_top, 2))
            else:
                self.lbl_card_barcode_img.pack_configure(pady=(0, max(0, 2 + pad_top)))

            self.barcode_preview_img = ctk_bc_img

    def gerar_etiquetas_word(self):
        if self.df_carregado is None:
            messagebox.showerror("Erro", "Selecione e carregue uma planilha primeiro!", parent=self)
            return

        caminho_salvar = filedialog.asksaveasfilename(
            defaultextension=".docx", filetypes=[("Documento Word", "*.docx")], parent=self
        )
        if not caminho_salvar:
            return

        lista_etiquetas = []
        col_qtde = self.combo_qtde.get()
        qtde_manual_val = int(self.obter_valor_float(self.entry_qtde_manual, padrao=1))

        for _, row in self.df_carregado.iterrows():
            qtd = qtde_manual_val
            if col_qtde != "Nenhum" and pd.notna(row[col_qtde]):
                try:
                    qtd = int(row[col_qtde])
                except ValueError:
                    qtd = qtde_manual_val
            for _ in range(max(1, qtd)):
                lista_etiquetas.append(row)

        cfg_modelo = MODELOS_PIMACO.get(self.combo_modelo_pimaco.get(), MODELOS_PIMACO["Pimaco A4251 (65 etq - 3.82 x 2.12 cm)"])

        f_top = self.obter_valor_float(self.entry_fonte_top, padrao=5.5)
        f_desc = self.obter_valor_float(self.entry_fonte_desc, padrao=4.5)
        entrelinha_val = self.obter_valor_float(self.entry_entrelinha, padrao=0.85)
        espaco_topo_desc_val = self.obter_valor_float(self.entry_espaco_topo_desc, padrao=0.2)
        pct_num = self.obter_valor_float(self.entry_fonte_numeral, padrao=30.0)
        largura_barras_cm = self.obter_valor_float(self.entry_largura_barras, padrao=cfg_modelo["width_cm"])
        altura_barras_mm = self.obter_valor_float(self.entry_altura_barras, padrao=1.8)
        distancia_numeral_val = self.obter_valor_float(self.entry_distancia_numeral, padrao=0.1)
        espaco_antes_bc = self.obter_valor_float(self.entry_espaco_antes_bc, padrao=0.0)
        pad_top_val = self.obter_valor_float(self.entry_pad_top, padrao=0.0)
        pad_left_val = self.obter_valor_float(self.entry_pad_left, padrao=0.0)

        limite_max_cm = cfg_modelo["width_cm"] - 0.2
        if largura_barras_cm <= 0 or largura_barras_cm > limite_max_cm:
            largura_barras_cm = limite_max_cm

        f_num_pt = max(4.0, (pct_num / 100.0) * 20.0)

        try:
            doc = docx.Document()
            section = doc.sections[0]

            section.page_width = Cm(21.0)
            section.page_height = Cm(29.7)
            section.top_margin = Cm(cfg_modelo["top_m"])
            section.bottom_margin = Cm(cfg_modelo["bot_m"])
            section.left_margin = Cm(cfg_modelo["left_m"])
            section.right_margin = Cm(cfg_modelo["right_m"])

            posicao_inicial = int(self.obter_valor_float(self.entry_posicao_inicial, padrao=1))
            offset = max(0, posicao_inicial - 1)

            cols_num = cfg_modelo["cols"]
            col_width_dxa = int(cfg_modelo["width_cm"] * 567)

            tabela = doc.add_table(rows=0, cols=cols_num)
            tabela.autofit = False
            tabela.alignment = WD_TABLE_ALIGNMENT.CENTER

            tblGrid = parse_xml(
                f'<w:tblGrid {nsdecls("w")}>' +
                ''.join([f'<w:gridCol w:w="{col_width_dxa}"/>' for _ in range(cols_num)]) +
                '</w:tblGrid>'
            )
            tabela._tbl.append(tblGrid)

            aplicar_bordas_tabela(tabela, exibir_bordas=(self.chk_incluir_bordas.get() == 1))

            total_itens = len(lista_etiquetas) + offset
            total_linhas = (total_itens + cols_num - 1) // cols_num

            for r_idx in range(total_linhas):
                row = tabela.add_row()
                trPr = row._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                # Ajustado para atLeast para prevenir achatamento na altura da célula
                trPr.append(parse_xml(f'<w:trHeight {nsdecls("w")} w:val="{cfg_modelo["row_height"]}" w:hRule="atLeast"/>'))

            index_etiqueta = 0
            for idx_celula in range(total_itens):
                r_idx = idx_celula // cols_num
                c_idx = idx_celula % cols_num
                cell = tabela.cell(r_idx, c_idx)
                cell.width = Cm(cfg_modelo["width_cm"])
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                if idx_celula < offset:
                    cell.text = ""
                    continue

                if index_etiqueta >= len(lista_etiquetas):
                    cell.text = ""
                    continue

                row_data = lista_etiquetas[index_etiqueta]
                index_etiqueta += 1

                def pegar_val_str(combo):
                    c = combo.get()
                    if c != "Nenhum" and c in row_data and pd.notna(row_data[c]):
                        v = str(row_data[c]).strip()
                        return v[:-2] if v.endswith(".0") else v
                    return ""

                v_cod = pegar_val_str(self.combo_codigo)
                v_rua = pegar_val_str(self.combo_rua)
                v_lin = pegar_val_str(self.combo_linha)
                v_dsc = pegar_val_str(self.combo_desc)
                v_ean = pegar_val_str(self.combo_ean)

                cell.text = ""
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER

                aplicar_espacamento_paragrafo(
                    p,
                    space_before_pt=pad_top_val,
                    space_after_pt=espaco_topo_desc_val,
                    line_spacing=entrelinha_val
                )
                aplicar_recuo_esquerda(p, pad_left_val)

                topo_parts = []
                if v_cod: topo_parts.append(v_cod)
                if v_rua: topo_parts.append(f"({v_rua})")
                if v_lin: topo_parts.append(f"({v_lin})")
                str_topo = " ".join(topo_parts)

                if str_topo:
                    run_top = p.add_run(str_topo)
                    run_top.font.name = "Arial"
                    run_top.font.size = Pt(f_top)
                    run_top.font.bold = True

                if v_dsc:
                    p_desc = cell.add_paragraph()
                    p_desc.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    aplicar_espacamento_paragrafo(
                        p_desc,
                        space_before_pt=0,
                        space_after_pt=espaco_antes_bc,
                        line_spacing=entrelinha_val
                    )
                    aplicar_recuo_esquerda(p_desc, pad_left_val)

                    run_dsc = p_desc.add_run(v_dsc)
                    run_dsc.font.name = "Arial"
                    run_dsc.font.size = Pt(f_desc)

                val_ean_limpo = remover_acentos(v_ean).strip() if v_ean else ""
                if val_ean_limpo:
                    tipo_code = self.combo_tipo_codigo.get()
                    caminho_img = gerar_imagem_codigo(
                        tipo=tipo_code,
                        codigo_str=val_ean_limpo,
                        font_size_pt=f_num_pt,
                        module_height_mm=altura_barras_mm,
                        text_distance=distancia_numeral_val,
                    )

                    if caminho_img and os.path.exists(caminho_img):
                        p_bc = cell.add_paragraph()
                        p_bc.alignment = WD_ALIGN_PARAGRAPH.CENTER

                        # Entrelinha simples (1.0) para garantir que a imagem não sobreponha os textos
                        aplicar_espacamento_paragrafo(
                            p_bc,
                            space_before_pt=max(0, espaco_antes_bc),
                            space_after_pt=0,
                            line_spacing=1.0
                        )
                        aplicar_recuo_esquerda(p_bc, pad_left_val)

                        run_bc = p_bc.add_run()
                        run_bc.add_picture(caminho_img, width=Cm(largura_barras_cm))

            doc.save(caminho_salvar)
            messagebox.showinfo("Sucesso", "Documento Word gerado com sucesso!", parent=self)

        except Exception as e:
            messagebox.showerror("Erro ao Gerar", f"Ocorreu uma falha ao gerar o arquivo Word:\n{e}", parent=self)


if __name__ == "__main__":
    app = AppEtiquetas()
    app.mainloop()