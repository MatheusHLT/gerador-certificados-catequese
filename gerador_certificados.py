from pathlib import Path
import calendar
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import date, datetime

from docx import Document
from docxcompose.composer import Composer

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).resolve().parent

TEMPLATE = APP_DIR / "modelo_certificado.docx"
OUTPUT_DIR = APP_DIR / "certificados_gerados"

MESES = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
]

CIDADE_PADRAO = "Belo Horizonte"

# ----------------------------------------------------------------------------
# Datas
# ----------------------------------------------------------------------------

def parse_data(valor: str) -> date:
    valor = valor.strip()
    for formato in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(valor, formato).date()
        except ValueError:
            pass
    raise ValueError("Informe a data no formato DD/MM/AAAA (ex.: 22/11/2026).")


def extenso(data: date) -> str:
    return f"{data.day} de {MESES[data.month - 1]} de {data.year}"


def data_por_extenso(valor: str) -> str:
    return extenso(parse_data(valor))

# ----------------------------------------------------------------------------
# Leitura da lista de casais
# ----------------------------------------------------------------------------

def ler_casais(caminho: Path):
    casais = []
    for linha in caminho.read_text(encoding="utf-8-sig").splitlines():
        linha = linha.strip()
        if not linha:
            continue
        # Aceita linhas como '1. Nome A e Nome B' ou apenas 'Nome A e Nome B'.
        linha = re.sub(r"^\s*\d+\s*[\.\)\-]\s*", "", linha).strip()
        if " e " not in linha:
            raise ValueError(
                f"Não consegui identificar os dois nomes nesta linha:\n{linha}\n\n"
                "Use o formato: Nome do noivo e Nome da noiva."
            )
        noivo, noiva = linha.split(" e ", 1)
        noivo, noiva = noivo.strip(), noiva.strip()
        if not noivo or not noiva:
            raise ValueError(f"Nome incompleto nesta linha:\n{linha}")
        casais.append((noivo, noiva))
    if not casais:
        raise ValueError("O arquivo TXT não contém nenhum casal.")
    return casais

# ----------------------------------------------------------------------------
# Preenchimento do modelo Word
# ----------------------------------------------------------------------------

# Reconhece os três campos do modelo sem depender do texto exato, para que
# pequenas edições no certificado (igreja, padre, data de exemplo) não o quebrem.
RE_CAMPO_NOMES = re.compile(r"nome\s+do\s+noivo", re.IGNORECASE)
RE_ENCONTRO = re.compile(r"(\bno\s+dia\s+)(.+?)(\s*,\s*realizado\b)", re.IGNORECASE | re.DOTALL)
RE_DATA_EXTENSO = re.compile(r"\d{1,2}\s+de\s+[A-Za-zÀ-ÿ]+\s+de\s+\d{4}")
# Ex.: 'Belo Horizonte, 28 de novembro de 2026.' (também aceita '99 de novembro...')
RE_EMISSAO = re.compile(r"\s*[^,]+,\s*\S+\s+de\s+\S+\s+de\s+\S+\s*")


def substituir_paragrafo(paragraph, novo_texto):
    # Mantém a formatação do primeiro trecho do parágrafo do modelo.
    if not paragraph.runs:
        paragraph.add_run(novo_texto)
        return
    paragraph.runs[0].text = novo_texto
    for run in paragraph.runs[1:]:
        run.text = ""


def texto_com_nova_data(texto, data_encontro):
    """Troca só a data do encontro e preserva o resto do parágrafo
    (inclusive o nome da igreja que estiver escrito no modelo)."""
    m = RE_ENCONTRO.search(texto)
    if m:
        return texto[:m.start(2)] + data_encontro + texto[m.end(2):]
    m = RE_DATA_EXTENSO.search(texto)
    if m:
        return texto[:m.start()] + data_encontro + texto[m.end():]
    return None


def atualizar_certificado(doc, noivo, noiva, data_encontro, emissao):
    encontrou_nomes = encontrou_encontro = encontrou_emissao = False
    for p in doc.paragraphs:
        texto = p.text
        if not texto.strip():
            continue
        if not encontrou_nomes and RE_CAMPO_NOMES.search(texto):
            substituir_paragrafo(p, f"{noivo} e {noiva}")
            encontrou_nomes = True
        elif not encontrou_encontro and "participaram" in texto.lower():
            novo = texto_com_nova_data(texto, data_encontro)
            if novo is not None:
                substituir_paragrafo(p, novo)
                encontrou_encontro = True
        elif not encontrou_emissao and RE_EMISSAO.fullmatch(texto):
            substituir_paragrafo(p, emissao)
            encontrou_emissao = True

    if not (encontrou_nomes and encontrou_encontro and encontrou_emissao):
        faltas = []
        if not encontrou_nomes:
            faltas.append("campo dos nomes (texto 'Nome do Noivo e Nome da Noiva')")
        if not encontrou_encontro:
            faltas.append("data do encontro (parágrafo 'participaram da preparação...')")
        if not encontrou_emissao:
            faltas.append("cidade/data de emissão (ex.: 'Belo Horizonte, 28 de novembro de 2026.')")
        raise ValueError(
            "Não encontrei no modelo: " + "; ".join(faltas) +
            ".\n\nVerifique se esses textos continuam no modelo_certificado.docx."
        )


def gerar(casais, data_encontro, emissao, destino):
    modelo = Document(str(TEMPLATE))
    atualizar_certificado(modelo, casais[0][0], casais[0][1], data_encontro, emissao)
    composer = Composer(modelo)

    for noivo, noiva in casais[1:]:
        doc = Document(str(TEMPLATE))
        atualizar_certificado(doc, noivo, noiva, data_encontro, emissao)
        composer.doc.add_page_break()
        composer.append(doc)

    destino.parent.mkdir(parents=True, exist_ok=True)
    composer.save(str(destino))

# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------

COR_TOPO = "#1f3a5f"
COR_FUNDO = "#f4f6f9"
COR_DESTAQUE = "#0078d7"
FONTE = ("Segoe UI", 10)
FONTE_NEGRITO = ("Segoe UI", 10, "bold")


class Calendario(tk.Toplevel):
    """Janela 'Escolher data': navegação por mês, dias clicáveis, Limpar e Hoje."""

    def __init__(self, parent, ancora, inicial, ao_selecionar, ao_limpar):
        super().__init__(parent)
        self.title("Escolher data")
        self.resizable(False, False)
        self.configure(bg="white")
        self.transient(parent.winfo_toplevel())

        self._ao_selecionar = ao_selecionar
        self._ao_limpar = ao_limpar
        self.hoje = date.today()
        self.selecionada = inicial
        referencia = inicial or self.hoje
        self.ano, self.mes = referencia.year, referencia.month

        self._montar()
        self._posicionar(ancora)
        self.bind("<Escape>", lambda _e: self.destroy())
        try:
            self.wait_visibility()
            self.grab_set()
        except tk.TclError:
            pass
        self.focus_set()

    def _montar(self):
        topo = tk.Frame(self, bg="white")
        topo.pack(fill="x", padx=12, pady=(12, 6))
        tk.Button(topo, text="◄", width=3, relief="solid", bd=1, bg="white",
                  command=lambda: self._mudar_mes(-1)).pack(side="left")
        tk.Button(topo, text="►", width=3, relief="solid", bd=1, bg="white",
                  command=lambda: self._mudar_mes(1)).pack(side="right")
        self.lbl_mes = tk.Label(topo, bg="white", font=("Segoe UI", 12, "bold"))
        self.lbl_mes.pack(expand=True)

        self.frame_dias = tk.Frame(self, bg="white")
        self.frame_dias.pack(padx=12, pady=4)

        rodape = tk.Frame(self, bg="white")
        rodape.pack(fill="x", padx=12, pady=(8, 12))
        tk.Button(rodape, text="Limpar", width=10, relief="solid", bd=1, bg="white",
                  command=self._limpar).pack(side="left")
        tk.Button(rodape, text="Hoje", width=10, relief="solid", bd=1, bg="white",
                  command=lambda: self._escolher(self.hoje)).pack(side="right")

        self._desenhar()

    def _desenhar(self):
        self.lbl_mes.config(text=f"{MESES[self.mes - 1]} de {self.ano}")
        for w in self.frame_dias.winfo_children():
            w.destroy()
        for c, letra in enumerate("DSTQQSS"):
            tk.Label(self.frame_dias, text=letra, width=4, bg="white",
                     font=FONTE_NEGRITO).grid(row=0, column=c, pady=(0, 4))
        semanas = calendar.Calendar(firstweekday=6).monthdatescalendar(self.ano, self.mes)
        for r, semana in enumerate(semanas, start=1):
            for c, dia in enumerate(semana):
                fora_do_mes = dia.month != self.mes
                selecionado = dia == self.selecionada
                fg = "#9a9a9a" if fora_do_mes else "#111111"
                bg = "white"
                if selecionado:
                    fg, bg = "white", COR_DESTAQUE
                fonte = FONTE_NEGRITO if dia == self.hoje else FONTE
                tk.Button(
                    self.frame_dias, text=str(dia.day), width=4, relief="flat", bd=0,
                    bg=bg, fg=fg, font=fonte, cursor="hand2",
                    activebackground=bg if selecionado else "#cce4f7",
                    activeforeground=fg,
                    command=lambda d=dia: self._escolher(d),
                ).grid(row=r, column=c, padx=1, pady=1)

    def _mudar_mes(self, delta):
        total = (self.mes - 1) + delta
        self.ano += total // 12
        self.mes = total % 12 + 1
        self._desenhar()

    def _escolher(self, dia):
        self._ao_selecionar(dia)
        self.destroy()

    def _limpar(self):
        self._ao_limpar()
        self.destroy()

    def _posicionar(self, ancora):
        self.update_idletasks()
        largura, altura = self.winfo_reqwidth(), self.winfo_reqheight()
        x = ancora.winfo_rootx()
        y = ancora.winfo_rooty() + ancora.winfo_height() + 2
        tela_l, tela_a = self.winfo_screenwidth(), self.winfo_screenheight()
        x = max(0, min(x, tela_l - largura - 10))
        if y + altura > tela_a - 40:
            y = max(0, ancora.winfo_rooty() - altura - 2)
        self.geometry(f"+{x}+{y}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gerador de Certificados — Catequese Matrimonial")
        self.configure(bg=COR_FUNDO)
        self.resizable(False, False)

        estilo = ttk.Style(self)
        estilo.theme_use("vista" if "vista" in estilo.theme_names() else "clam")
        estilo.configure("Gerar.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 8))
        estilo.configure("TButton", font=FONTE, padding=(10, 5))

        hoje = date.today()
        self.var_data = tk.StringVar(value=hoje.strftime("%d/%m/%Y"))
        self.var_emissao = tk.StringVar(value=f"{CIDADE_PADRAO}, {extenso(hoje)}")
        self.var_arquivo = tk.StringVar()
        self.var_status = tk.StringVar()
        self.arquivo_txt = None

        self._montar()
        self._centralizar(680, 560)

        if not TEMPLATE.exists():
            self.var_status.set("⚠ Modelo não encontrado: coloque 'modelo_certificado.docx' "
                                "na mesma pasta do programa.")
            self.btn_gerar.state(["disabled"])
            messagebox.showerror(
                "Modelo não encontrado",
                "Coloque o arquivo 'modelo_certificado.docx' na mesma pasta deste programa.",
                parent=self,
            )

    # -- construção da tela --------------------------------------------------
    def _montar(self):
        topo = tk.Frame(self, bg=COR_TOPO)
        topo.pack(fill="x")
        tk.Label(topo, text="Gerador de Certificados", bg=COR_TOPO, fg="white",
                 font=("Segoe UI", 22, "bold")).pack(pady=(22, 0))
        tk.Label(topo, text="Catequese Matrimonial", bg=COR_TOPO, fg="#c9d6e8",
                 font=("Segoe UI", 12)).pack(pady=(0, 20))

        corpo = tk.Frame(self, bg=COR_FUNDO)
        corpo.pack(fill="both", expand=True, padx=36, pady=(22, 0))
        corpo.columnconfigure(0, weight=1)

        # 1. Data do encontro
        self._rotulo(corpo, "1.  Data do encontro", 0)
        linha = tk.Frame(corpo, bg=COR_FUNDO)
        linha.grid(row=1, column=0, sticky="w")
        self.entry_data = ttk.Entry(linha, textvariable=self.var_data, width=16,
                                    font=("Segoe UI", 11))
        self.entry_data.pack(side="left", ipady=3)
        self.btn_calendario = ttk.Button(linha, text="📅", width=4,
                                         command=self._abrir_calendario)
        self.btn_calendario.pack(side="left", padx=(6, 0))
        self._dica(corpo, "Digite no formato DD/MM/AAAA ou escolha no calendário.", 2)

        # 2. Cidade e data de emissão
        self._rotulo(corpo, "2.  Cidade e data de emissão", 3)
        ttk.Entry(corpo, textvariable=self.var_emissao,
                  font=("Segoe UI", 11)).grid(row=4, column=0, sticky="ew", ipady=3)
        self._dica(corpo, "Como deve aparecer no certificado. Ex.: Belo Horizonte, "
                          "21 de outubro de 2026", 5)

        # 3. Lista de casais
        self._rotulo(corpo, "3.  Lista de casais (.txt)", 6)
        linha = tk.Frame(corpo, bg=COR_FUNDO)
        linha.grid(row=7, column=0, sticky="ew")
        linha.columnconfigure(0, weight=1)
        ttk.Entry(linha, textvariable=self.var_arquivo, state="readonly",
                  font=("Segoe UI", 10)).grid(row=0, column=0, sticky="ew", ipady=3)
        ttk.Button(linha, text="Procurar...", command=self._escolher_arquivo).grid(
            row=0, column=1, padx=(6, 0))
        self._dica(corpo, "Uma linha por casal: Nome do noivo e Nome da noiva "
                          "(a numeração é opcional).", 8)

        # Rodapé
        rodape = tk.Frame(self, bg=COR_FUNDO)
        rodape.pack(fill="x", padx=36, pady=(14, 22))
        tk.Label(rodape, textvariable=self.var_status, bg=COR_FUNDO, fg="#b00020",
                 font=FONTE, wraplength=600, justify="left").pack(anchor="w", pady=(0, 8))
        botoes = tk.Frame(rodape, bg=COR_FUNDO)
        botoes.pack(fill="x")
        self.btn_gerar = ttk.Button(botoes, text="Gerar certificados",
                                    style="Gerar.TButton", command=self._gerar)
        self.btn_gerar.pack(side="left")
        ttk.Button(botoes, text="Sair", command=self.destroy).pack(side="right")

    def _rotulo(self, pai, texto, linha):
        tk.Label(pai, text=texto, bg=COR_FUNDO, fg="#1b1b1b",
                 font=("Segoe UI", 11, "bold")).grid(
            row=linha, column=0, sticky="w", pady=(14 if linha else 0, 4))

    def _dica(self, pai, texto, linha):
        tk.Label(pai, text=texto, bg=COR_FUNDO, fg="#666666",
                 font=("Segoe UI", 9)).grid(row=linha, column=0, sticky="w", pady=(3, 0))

    def _centralizar(self, largura, altura):
        self.update_idletasks()
        x = (self.winfo_screenwidth() - largura) // 2
        y = max(0, (self.winfo_screenheight() - altura) // 3)
        self.geometry(f"{largura}x{altura}+{x}+{y}")

    # -- ações ---------------------------------------------------------------
    def _abrir_calendario(self):
        try:
            inicial = parse_data(self.var_data.get())
        except ValueError:
            inicial = None
        Calendario(
            self, self.entry_data, inicial,
            ao_selecionar=lambda d: self.var_data.set(d.strftime("%d/%m/%Y")),
            ao_limpar=lambda: self.var_data.set(""),
        )

    def _escolher_arquivo(self):
        caminho = filedialog.askopenfilename(
            parent=self,
            title="Selecione a lista de casais (.txt)",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")],
        )
        if caminho:
            self.arquivo_txt = Path(caminho)
            self.var_arquivo.set(caminho)

    def _gerar(self):
        self.var_status.set("")
        try:
            data_encontro = data_por_extenso(self.var_data.get())
        except ValueError as exc:
            messagebox.showerror("Data inválida", str(exc), parent=self)
            self.entry_data.focus_set()
            return

        emissao = self.var_emissao.get().strip()
        if not emissao:
            messagebox.showwarning("Campo vazio",
                                   "Informe a cidade e a data de emissão.", parent=self)
            return

        if not self.arquivo_txt:
            messagebox.showwarning("Lista de casais",
                                   "Selecione o arquivo .txt com a lista de casais.",
                                   parent=self)
            return

        self.btn_gerar.state(["disabled"])
        self.var_status.set("Gerando certificados, aguarde...")
        self.update_idletasks()
        try:
            casais = ler_casais(self.arquivo_txt)
            nome_saida = f"Certificados_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            destino = OUTPUT_DIR / nome_saida
            gerar(casais, data_encontro, emissao, destino)
        except Exception as exc:
            self.var_status.set("")
            messagebox.showerror("Não foi possível gerar", str(exc), parent=self)
            return
        finally:
            self.btn_gerar.state(["!disabled"])

        self.var_status.set("")
        messagebox.showinfo(
            "Certificados gerados",
            f"Foram gerados {len(casais)} certificados, um por página.\n\n"
            f"Arquivo salvo em:\n{destino}",
            parent=self,
        )


def main():
    App().mainloop()


if __name__ == "__main__":
    main()
