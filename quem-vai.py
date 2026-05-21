#!/usr/bin/env python3
"""Quem Vai? — Ferramenta de sorteio para o treinador Luis Reis."""

import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk
import sqlite3
import random
import os
import sys
import json
from tkinter import messagebox

# ─── Tema ─────────────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

F_TITLE  = ("Segoe UI", 26, "bold")
F_LARGE  = ("Segoe UI", 20, "bold")
F_MEDIUM = ("Segoe UI", 16)
F_SMALL  = ("Segoe UI", 14)
BTN_H    = 48
NAV_ACTIVE   = ("#1d4ed8", "#1d4ed8")
NAV_INACTIVE = "transparent"

# ─── Estilos ttk (Treeview Excel-like) ────────────────────────────────────────
_STYLE_DONE = False

def _setup_treeview_style():
    global _STYLE_DONE
    if _STYLE_DONE:
        return
    _STYLE_DONE = True
    s = ttk.Style()
    try:
        s.theme_use("clam")
    except Exception:
        pass
    s.configure("Excel.Treeview",
        background="#1e1e2e",
        foreground="#ffffff",
        fieldbackground="#1e1e2e",
        bordercolor="#2a2a4a",
        borderwidth=0,
        rowheight=34,
        font=("Segoe UI", 13),
    )
    s.configure("Excel.Treeview.Heading",
        background="#28283a",
        foreground="#8899bb",
        font=("Segoe UI", 11, "bold"),
        relief="flat",
        borderwidth=0,
        padding=(8, 5),
    )
    s.map("Excel.Treeview",
        background=[("selected", "#1d4ed8")],
        foreground=[("selected", "#ffffff")],
    )
    s.map("Excel.Treeview.Heading",
        background=[("active", "#1a2a4a")],
    )
    s.configure("Excel.Vertical.TScrollbar",
        background="#2a2a4a",
        troughcolor="#16213e",
        arrowcolor="#8899bb",
        borderwidth=0,
    )


# ─── Base de Dados ────────────────────────────────────────────────────────────

def _app_folder() -> str:
    base = os.environ.get("APPDATA") or os.path.dirname(
        sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    )
    return os.path.join(base, "quem-vai")


def _db_path() -> str:
    folder = _app_folder()
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "quem-vai.db")


def _config_path() -> str:
    return os.path.join(_app_folder(), "config.json")


def _load_config() -> dict:
    try:
        with open(_config_path(), encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_config(cfg: dict) -> None:
    os.makedirs(_app_folder(), exist_ok=True)
    with open(_config_path(), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False)


def _conn() -> sqlite3.Connection:
    c = sqlite3.connect(_db_path())
    c.row_factory = lambda cur, row: {
        col[0]: row[i] for i, col in enumerate(cur.description)
    }
    c.execute("PRAGMA foreign_keys = ON")
    return c


def _init_db() -> None:
    with _conn() as c:
        c.executescript("""
            CREATE TABLE IF NOT EXISTS equipas (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE,
                tipo TEXT NOT NULL DEFAULT 'CEF' CHECK(tipo IN ('CEF','CP'))
            );
            CREATE TABLE IF NOT EXISTS jogadores (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                nome           TEXT NOT NULL,
                tipo           TEXT NOT NULL CHECK(tipo IN ('CEF','CP')),
                numero_interno TEXT,
                equipa_id      INTEGER NOT NULL REFERENCES equipas(id) ON DELETE CASCADE
            );
        """)
    with _conn() as c:
        try:
            c.execute("ALTER TABLE jogadores ADD COLUMN numero_interno TEXT")
        except sqlite3.OperationalError:
            pass
    with _conn() as c:
        try:
            c.execute("ALTER TABLE equipas ADD COLUMN tipo TEXT NOT NULL DEFAULT 'CEF' CHECK(tipo IN ('CEF','CP'))")
        except sqlite3.OperationalError:
            pass


# ─── Helpers DB ───────────────────────────────────────────────────────────────

def listar_equipas() -> list[dict]:
    with _conn() as c:
        return c.execute(
            "SELECT e.id, e.nome, e.tipo, COUNT(j.id) AS total"
            " FROM equipas e LEFT JOIN jogadores j ON j.equipa_id=e.id"
            " GROUP BY e.id ORDER BY e.nome"
        ).fetchall()


def listar_alunos(equipa_id: int) -> list[dict]:
    with _conn() as c:
        return c.execute(
            "SELECT id, nome, numero_interno"
            " FROM jogadores WHERE equipa_id=? ORDER BY nome",
            (equipa_id,)
        ).fetchall()


def criar_equipa(nome: str, tipo: str) -> None:
    with _conn() as c:
        c.execute("INSERT INTO equipas (nome, tipo) VALUES (?,?)", (nome, tipo))


def atualizar_equipa(eid: int, novo_nome: str, novo_tipo: str) -> None:
    with _conn() as c:
        c.execute("UPDATE equipas SET nome=?, tipo=? WHERE id=?", (novo_nome, novo_tipo, eid))


def apagar_equipa(eid: int) -> None:
    with _conn() as c:
        c.execute("DELETE FROM equipas WHERE id=?", (eid,))


def adicionar_aluno(nome: str, numero_interno: str, equipa_id: int) -> None:
    with _conn() as c:
        eq = c.execute("SELECT tipo FROM equipas WHERE id=?", (equipa_id,)).fetchone()
        tipo = eq["tipo"] if eq else "CEF"
        c.execute(
            "INSERT INTO jogadores (nome, tipo, numero_interno, equipa_id) VALUES (?,?,?,?)",
            (nome, tipo, numero_interno or None, equipa_id)
        )


def apagar_aluno(jid: int) -> None:
    with _conn() as c:
        c.execute("DELETE FROM jogadores WHERE id=?", (jid,))


# ─── Lógica de Sorteio ────────────────────────────────────────────────────────

class Sorteio:
    def __init__(self):
        self.filas: dict[int, list[int]] = {}
        self.historico: list[dict] = []

    def iniciar(self) -> None:
        self.filas.clear()
        with _conn() as c:
            for eq in c.execute("SELECT id FROM equipas").fetchall():
                ids = [r["id"] for r in c.execute(
                    "SELECT id FROM jogadores WHERE equipa_id=?", (eq["id"],)
                ).fetchall()]
                random.shuffle(ids)
                self.filas[eq["id"]] = ids

    def sortear(self) -> dict | None:
        disponiveis = [
            (eid, jid) for eid, fila in self.filas.items() for jid in fila
        ]
        if not disponiveis:
            self.iniciar()
            disponiveis = [
                (eid, jid) for eid, fila in self.filas.items() for jid in fila
            ]
        if not disponiveis:
            return None

        eid, jid = random.choice(disponiveis)
        self.filas[eid].remove(jid)

        if not self.filas[eid]:
            with _conn() as c:
                ids = [r["id"] for r in c.execute(
                    "SELECT id FROM jogadores WHERE equipa_id=?", (eid,)
                ).fetchall()]
                random.shuffle(ids)
                self.filas[eid] = ids

        with _conn() as c:
            row = c.execute(
                "SELECT j.nome, e.tipo, j.numero_interno, e.nome AS equipa"
                " FROM jogadores j JOIN equipas e ON e.id=j.equipa_id WHERE j.id=?",
                (jid,)
            ).fetchone()

        self.historico.append(row)
        return row


# ─── Diálogo de Primeiro Uso ─────────────────────────────────────────────────

class DialogoPrimeiroUso(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Bem-vindo!")
        self.geometry("420x260")
        self.resizable(False, False)
        self.nome = ""

        ctk.CTkLabel(self, text="Bem-vindo ao Quem Vai?", font=F_LARGE).pack(pady=(28, 6))
        ctk.CTkLabel(
            self,
            text="Para começar, qual é o teu nome?",
            font=F_MEDIUM, text_color="gray60"
        ).pack(pady=(0, 12))

        self.entry = ctk.CTkEntry(self, font=F_MEDIUM, height=BTN_H, width=340,
                                  placeholder_text="Ex: Luis Reis")
        self.entry.pack()

        ctk.CTkButton(self, text="Começar", font=F_MEDIUM, height=BTN_H,
                      command=self._guardar).pack(pady=20)
        self.entry.bind("<Return>", lambda _: self._guardar())
        self.after(150, self._on_show)

    def _on_show(self):
        self.lift()
        self.grab_set()
        self.entry.focus_set()

    def _guardar(self):
        nome = self.entry.get().strip()
        if not nome:
            messagebox.showerror("Erro", "Escreve o teu nome para continuar.", parent=self)
            return
        self.nome = nome
        self.destroy()


# ─── Diálogo de Equipa ────────────────────────────────────────────────────────

class DialogoEquipa(ctk.CTkToplevel):
    def __init__(self, parent, equipa: dict | None = None):
        super().__init__(parent)
        self.title("Equipa")
        self.geometry("380x300")
        self.resizable(False, False)
        self.resultado = None

        ctk.CTkLabel(self, text="Nome da equipa:", font=F_MEDIUM).pack(pady=(20, 6))
        self.entry = ctk.CTkEntry(self, font=F_MEDIUM, height=BTN_H, width=320)
        self.entry.pack()
        if equipa:
            self.entry.insert(0, equipa["nome"])

        ctk.CTkLabel(self, text="Tipo de alunos:", font=F_MEDIUM).pack(pady=(14, 6))
        self._seg_tipo = ctk.CTkSegmentedButton(
            self, values=["CEF", "CP"],
            font=F_MEDIUM, width=320, height=BTN_H,
        )
        self._seg_tipo.set(equipa["tipo"] if equipa else "CEF")
        self._seg_tipo.pack()

        ctk.CTkButton(self, text="Guardar", font=F_MEDIUM, height=BTN_H,
                      command=self._guardar).pack(pady=14)
        self.entry.bind("<Return>", lambda _: self._guardar())
        self.after(150, self._on_show)

    def _on_show(self):
        self.lift()
        self.grab_set()
        self.entry.focus_set()

    def _guardar(self):
        nome = self.entry.get().strip()
        if not nome:
            messagebox.showerror("Erro", "O nome não pode estar vazio.", parent=self)
            return
        self.resultado = {"nome": nome, "tipo": self._seg_tipo.get()}
        self.destroy()


# ─── Tabela Excel ─────────────────────────────────────────────────────────────

class TabelaAlunos(tk.Frame):
    """Grid editável estilo Excel para os alunos de uma equipa."""

    NOVO_ID = "__novo__"

    def __init__(self, parent, equipa_id: int, on_change):
        super().__init__(parent, bg="#1e1e2e")
        self.equipa_id = equipa_id
        self.on_change = on_change
        self._popup: tk.Widget | None = None
        self._popup_item: str | None = None
        self._popup_col: int | None = None
        self._saving_novo = False
        self._build()

    def _build(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self,
            columns=("num", "nome"),
            show="headings",
            style="Excel.Treeview",
            selectmode="browse",
        )
        self.tree.heading("num",  text="Nº Interno", anchor="w")
        self.tree.heading("nome", text="Nome",        anchor="w")
        self.tree.column("num",  width=130, minwidth=80,  stretch=False)
        self.tree.column("nome", width=380, minwidth=120, stretch=True)

        vsb = ttk.Scrollbar(self, orient="vertical",
                            command=self.tree.yview,
                            style="Excel.Vertical.TScrollbar")
        self.tree.configure(yscrollcommand=vsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        # Rodapé com dica
        rodape = tk.Frame(self, bg="#28283a")
        rodape.grid(row=1, column=0, columnspan=2, sticky="ew")
        tk.Label(rodape,
                 text="Clica numa célula para editar  •  Tab/Enter para avançar  "
                      "•  Botão direito para apagar linha",
                 bg="#16213e", fg="#556677",
                 font=("Segoe UI", 10),
                 anchor="w", padx=8, pady=4
                 ).pack(fill="x")

        self.tree.tag_configure("alt",  background="#28283a", foreground="#ffffff")
        self.tree.tag_configure("novo", foreground="#445566",
                                font=("Segoe UI", 12, "italic"))

        self.tree.bind("<ButtonRelease-1>", self._on_click)
        self.tree.bind("<Return>", lambda e: self._on_enter())
        self.tree.bind("<Button-3>", self._on_right_click)

        self.load()

    # ── Carregamento ─────────────────────────────────────────────────────────

    def load(self):
        self._close_popup()
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, a in enumerate(listar_alunos(self.equipa_id)):
            tags = ("alt",) if i % 2 else ()
            self.tree.insert("", "end", iid=str(a["id"]),
                             values=(a["numero_interno"] or "", a["nome"]),
                             tags=tags)
        self.tree.insert("", "end", iid=self.NOVO_ID,
                         values=("", "← escreve aqui para adicionar aluno"),
                         tags=("novo",))

    # ── Interação ────────────────────────────────────────────────────────────

    def _on_click(self, event):
        item = self.tree.identify_row(event.y)
        col  = self.tree.identify_column(event.x)
        if item and col:
            self._open_popup(item, col)

    def _on_enter(self):
        item = self.tree.focus()
        if item:
            self._open_popup(item, "#1")

    def _on_right_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item or item == self.NOVO_ID:
            return
        self.tree.selection_set(item)
        menu = tk.Menu(self, tearoff=0,
                       bg="#1e1e2e", fg="#e0e0e0",
                       activebackground="#1d4ed8", activeforeground="white",
                       bd=0, relief="flat")
        menu.add_command(label="Apagar linha",
                         command=lambda: self._delete_row(int(item)))
        menu.post(event.x_root, event.y_root)

    # ── Popup de edição ──────────────────────────────────────────────────────

    def _open_popup(self, item: str, col: str):
        self._close_popup()

        bbox = self.tree.bbox(item, col)
        if not bbox:
            return

        x, y, w, h = bbox
        col_idx = int(col[1:]) - 1  # "#1" → 0

        vals = list(self.tree.item(item, "values"))
        while len(vals) < 2:
            vals.append("")
        cur = vals[col_idx]

        # Limpa placeholder da linha nova
        if item == self.NOVO_ID and col_idx == 1:
            cur = ""

        widget = tk.Entry(
            self.tree,
            font=("Segoe UI", 13),
            bg="#0d1b2a", fg="white",
            insertbackground="white",
            relief="flat", bd=0,
            highlightthickness=2,
            highlightcolor="#3b82f6",
            highlightbackground="#2a2a4a",
        )
        widget.insert(0, cur)
        widget.select_range(0, "end")

        widget.bind("<Return>",
                    lambda e, w=widget, i=item, c=col_idx: self._commit_next(i, c, w))
        widget.bind("<Tab>",
                    lambda e, w=widget, i=item, c=col_idx: (
                        self._commit_next(i, c, w), "break")[1])
        widget.bind("<Escape>",   lambda e: self._close_popup())
        widget.bind("<FocusOut>",
                    lambda e, w=widget, i=item, c=col_idx: self._commit(i, c, self._safe_get(w)))
        widget.bind("<Control-v>",
                    lambda e, w=widget, i=item, c=col_idx: self._paste_handler(e, i, c, w))

        widget.place(x=x, y=y, width=w, height=h)
        widget.focus_set()

        self._popup      = widget
        self._popup_item = item
        self._popup_col  = col_idx

    def _paste_handler(self, event, item: str, col_idx: int, widget):
        try:
            text = widget.clipboard_get()
        except Exception:
            return
        lines = [l for l in text.splitlines() if l.strip()]
        if len(lines) <= 1:
            return  # paste normal de uma linha — deixa o tkinter tratar

        self._close_popup()
        col_names = ["numero_interno", "nome"]
        all_items = [i for i in self.tree.get_children() if i != self.NOVO_ID]

        start_idx = len(all_items) if item == self.NOVO_ID else (
            all_items.index(item) if item in all_items else len(all_items)
        )

        for i, line in enumerate(lines):
            parts = line.split("\t")
            row_idx = start_idx + i

            if row_idx < len(all_items):
                existing = all_items[row_idx]
                vals = list(self.tree.item(existing, "values"))
                while len(vals) < 2:
                    vals.append("")
                for j, part in enumerate(parts):
                    tc = col_idx + j
                    if tc >= 2:
                        break
                    val = part.strip()
                    vals[tc] = val
                    try:
                        with _conn() as c:
                            c.execute(f"UPDATE jogadores SET {col_names[tc]}=? WHERE id=?",
                                      (val or None, int(existing)))
                    except Exception:
                        pass
                self.tree.item(existing, values=vals)
            else:
                num, nome = "", ""
                for j, part in enumerate(parts):
                    tc = col_idx + j
                    val = part.strip()
                    if tc == 0:
                        num = val
                    elif tc == 1:
                        nome = val
                if not nome and col_idx == 1 and len(parts) == 1:
                    nome = parts[0].strip()
                if nome:
                    adicionar_aluno(nome, num, self.equipa_id)

        self.load()
        self.on_change()
        return "break"

    def _close_popup(self):
        if self._popup is None:
            return
        popup = self._popup
        self._popup      = None  # limpa ANTES do destroy para evitar re-entrada via FocusOut
        self._popup_item = None
        self._popup_col  = None
        try:
            popup.destroy()
        except Exception:
            pass

    @staticmethod
    def _safe_get(widget) -> str | None:
        try:
            return widget.get()
        except Exception:
            return None

    # ── Commit ───────────────────────────────────────────────────────────────

    def _commit(self, item: str, col_idx: int, value: str | None):
        if self._popup is None or value is None:
            return
        self._close_popup()

        value = value.strip()

        vals = list(self.tree.item(item, "values"))
        while len(vals) < 2:
            vals.append("")
        vals[col_idx] = value

        if item == self.NOVO_ID:
            nome = vals[1]
            if nome and nome != "← escreve aqui para adicionar aluno":
                self._save_novo(vals[0], nome)
            return

        self.tree.item(item, values=vals)

        col_names = ["numero_interno", "nome"]
        col_name = col_names[col_idx]
        if col_name == "nome" and not value:
            return
        try:
            with _conn() as c:
                c.execute(f"UPDATE jogadores SET {col_name}=? WHERE id=?",
                          (value or None, int(item)))
            self.on_change()
        except Exception:
            pass

    def _commit_next(self, item: str, col_idx: int, widget):
        value = self._safe_get(widget)
        self._commit(item, col_idx, value)
        next_col = col_idx + 1
        if next_col >= 2:
            nxt = self.tree.next(item)
            if nxt:
                self.tree.selection_set(nxt)
                self.tree.focus(nxt)
                self._open_popup(nxt, "#1")
        else:
            self._open_popup(item, f"#{next_col + 1}")

    # ── Novo aluno ───────────────────────────────────────────────────────────

    def _save_novo(self, num: str, nome: str):
        if self._saving_novo:
            return
        self._saving_novo = True
        try:
            adicionar_aluno(nome, num, self.equipa_id)
            self.load()
            self.on_change()
        finally:
            self._saving_novo = False

    # ── Apagar ───────────────────────────────────────────────────────────────

    def _delete_row(self, aluno_id: int):
        with _conn() as c:
            row = c.execute(
                "SELECT nome FROM jogadores WHERE id=?", (aluno_id,)
            ).fetchone()
        if not row:
            return
        if not messagebox.askyesno("Confirmar", f"Apagar '{row['nome']}'?"):
            return
        apagar_aluno(aluno_id)
        self.load()
        self.on_change()


# ─── Frame: Sorteio ──────────────────────────────────────────────────────────

class SorteioFrame(ctk.CTkFrame):
    def __init__(self, parent, sorteio: Sorteio):
        super().__init__(parent, fg_color="transparent")
        self.sorteio = sorteio
        self._build()

    def _build(self):
        self.card = ctk.CTkFrame(self, corner_radius=12, height=170)
        self.card.pack(fill="x", padx=24, pady=(24, 12))
        self.card.pack_propagate(False)

        self.lbl_nome = ctk.CTkLabel(
            self.card, text="—", font=("Segoe UI", 36, "bold"), wraplength=580
        )
        self.lbl_nome.pack(expand=True, pady=(20, 4))

        self.lbl_info = ctk.CTkLabel(self.card, text="", font=F_MEDIUM, text_color="gray60")
        self.lbl_info.pack(pady=(0, 16))

        self.btn_sortear = ctk.CTkButton(
            self, text="SORTEAR", font=("Segoe UI", 24, "bold"),
            height=72, corner_radius=12, command=self._sortear
        )
        self.btn_sortear.pack(fill="x", padx=24, pady=8)

        ctk.CTkButton(
            self, text="Reiniciar sessão", font=F_SMALL, height=38,
            fg_color="transparent", border_width=1,
            hover_color=("gray85", "gray30"),
            command=self._reiniciar
        ).pack(padx=24, pady=(0, 16))

        ctk.CTkLabel(self, text="Histórico desta sessão", font=F_MEDIUM).pack(
            anchor="w", padx=24
        )
        self.lista = ctk.CTkScrollableFrame(self, height=180)
        self.lista.pack(fill="both", expand=True, padx=24, pady=(4, 24))

    def on_show(self):
        if not self.sorteio.historico:
            self.sorteio.iniciar()

    def _sortear(self):
        r = self.sorteio.sortear()
        if r is None:
            messagebox.showinfo(
                "Sem alunos",
                "Não há alunos configurados.\n"
                "Vai a 'Configurar' para adicionar equipas e alunos."
            )
            return
        self.lbl_nome.configure(text=r["nome"])
        num = f"  •  Nº {r['numero_interno']}" if r.get("numero_interno") else ""
        self.lbl_info.configure(text=f"{r['equipa']}  •  {r['tipo']}{num}")
        self._refresh_historico()

    def _reiniciar(self):
        self.sorteio.historico.clear()
        self.sorteio.iniciar()
        self.lbl_nome.configure(text="—")
        self.lbl_info.configure(text="")
        self._refresh_historico()

    def _refresh_historico(self):
        for w in self.lista.winfo_children():
            w.destroy()
        for r in reversed(self.sorteio.historico):
            num = f"  Nº {r['numero_interno']}" if r.get("numero_interno") else ""
            ctk.CTkLabel(
                self.lista,
                text=f"{r['nome']}{num}  —  {r['equipa']}  ({r['tipo']})",
                font=F_SMALL, anchor="w"
            ).pack(fill="x", padx=8, pady=2)


# ─── Frame: Configurar ───────────────────────────────────────────────────────

class ConfigFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self._equipa_sel: dict | None = None
        self._tabela: TabelaAlunos | None = None
        self._equipa_btns: dict[int, ctk.CTkButton] = {}
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ctk.CTkLabel(header, text="Configurar Equipas", font=F_LARGE).pack(side="left")
        ctk.CTkButton(
            header, text="+ Equipa", font=F_SMALL, height=40, width=130,
            command=self._nova_equipa
        ).pack(side="right")

        paineis = ctk.CTkFrame(self, fg_color="transparent")
        paineis.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        paineis.columnconfigure(0, weight=2)
        paineis.columnconfigure(1, weight=3)
        paineis.rowconfigure(0, weight=1)

        self.lista_equipas = ctk.CTkScrollableFrame(paineis, label_text="Equipas")
        self.lista_equipas.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.painel_dir = ctk.CTkFrame(paineis)
        self.painel_dir.grid(row=0, column=1, sticky="nsew")

        self._placeholder()

    def _placeholder(self):
        for w in self.painel_dir.winfo_children():
            w.destroy()
        self._tabela = None
        ctk.CTkLabel(
            self.painel_dir,
            text="Seleciona uma equipa para ver os alunos",
            font=F_MEDIUM, text_color="gray50"
        ).pack(expand=True)

    def on_show(self):
        self._refresh_equipas()

    # ── Equipas ──────────────────────────────────────────────────────────────

    def _refresh_equipas(self):
        self._equipa_btns.clear()
        for w in self.lista_equipas.winfo_children():
            w.destroy()
        for eq in listar_equipas():
            self._linha_equipa(eq)
        self._atualizar_cor_btns()

    def _linha_equipa(self, eq: dict):
        row = ctk.CTkFrame(self.lista_equipas, fg_color="transparent")
        row.pack(fill="x", pady=3)
        btn = ctk.CTkButton(
            row,
            text=f"{eq['nome']}\n{eq['tipo']} · {eq['total']} aluno{'s' if eq['total'] != 1 else ''}",
            font=F_SMALL, anchor="w", height=58,
            command=lambda e=eq: self._sel_equipa(e)
        )
        btn.pack(side="left", fill="x", expand=True)
        self._equipa_btns[eq["id"]] = btn
        ctk.CTkButton(row, text="✎", width=40, height=58, font=F_SMALL,
                      fg_color="transparent", border_width=1,
                      command=lambda e=eq: self._editar_equipa(e)).pack(side="left", padx=2)
        ctk.CTkButton(row, text="✕", width=40, height=58, font=F_SMALL,
                      fg_color="transparent", border_width=1,
                      text_color=("red", "#ef4444"),
                      command=lambda e=eq: self._apagar_equipa(e)).pack(side="left")

    def _nova_equipa(self):
        d = DialogoEquipa(self)
        self.wait_window(d)
        if not d.resultado:
            return
        try:
            criar_equipa(d.resultado["nome"], d.resultado["tipo"])
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", f"Já existe uma equipa chamada '{d.resultado['nome']}'.")
            return
        self._refresh_equipas()

    def _editar_equipa(self, eq: dict):
        d = DialogoEquipa(self, eq)
        self.wait_window(d)
        if not d.resultado:
            return
        try:
            atualizar_equipa(eq["id"], d.resultado["nome"], d.resultado["tipo"])
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", f"Já existe uma equipa chamada '{d.resultado['nome']}'.")
            return
        self._refresh_equipas()
        if self._equipa_sel and self._equipa_sel["id"] == eq["id"]:
            self._equipa_sel = {**self._equipa_sel, "nome": d.resultado["nome"], "tipo": d.resultado["tipo"]}
            self._refresh_tabela()

    def _apagar_equipa(self, eq: dict):
        if not messagebox.askyesno("Confirmar",
                                   f"Apagar '{eq['nome']}' e todos os seus alunos?"):
            return
        apagar_equipa(eq["id"])
        if self._equipa_sel and self._equipa_sel["id"] == eq["id"]:
            self._equipa_sel = None
            self._placeholder()
        self._refresh_equipas()

    def _atualizar_cor_btns(self):
        sel_id = self._equipa_sel["id"] if self._equipa_sel else None
        for eid, btn in self._equipa_btns.items():
            if eid == sel_id:
                btn.configure(fg_color=NAV_ACTIVE, hover_color=("#1e40af", "#1e40af"))
            else:
                btn.configure(fg_color=("#3b3b3b", "#3b3b3b"), hover_color=("gray85", "gray30"))

    def _sel_equipa(self, eq: dict):
        self._equipa_sel = eq
        self._atualizar_cor_btns()
        self._refresh_tabela()

    # ── Tabela de alunos ─────────────────────────────────────────────────────

    def _refresh_tabela(self):
        for w in self.painel_dir.winfo_children():
            w.destroy()
        self._tabela = None

        eq = self._equipa_sel
        if eq is None:
            self._placeholder()
            return

        topo = ctk.CTkFrame(self.painel_dir, fg_color="transparent")
        topo.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(topo, text=eq["nome"], font=F_LARGE).pack(side="left")

        self._tabela = TabelaAlunos(
            self.painel_dir,
            equipa_id=eq["id"],
            on_change=self._on_tabela_change,
        )
        self._tabela.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        equipas = {e["id"]: e for e in listar_equipas()}
        if eq["id"] in equipas:
            self._equipa_sel = equipas[eq["id"]]

    def _on_tabela_change(self):
        """Chamado pela tabela após qualquer modificação nos alunos."""
        self._refresh_equipas()
        equipas = {e["id"]: e for e in listar_equipas()}
        if self._equipa_sel and self._equipa_sel["id"] in equipas:
            self._equipa_sel = equipas[self._equipa_sel["id"]]


# ─── App Principal ────────────────────────────────────────────────────────────

class QuemVaiApp(ctk.CTk):
    def __init__(self, primeiro_uso: bool = False):
        super().__init__()
        self.title("Quem Vai?")
        self.geometry("940x660")
        self.minsize(720, 520)
        self.sorteio = Sorteio()
        self._build()
        self._nav("sorteio")
        if primeiro_uso:
            self.after(300, self._setup_primeiro_uso)

    def _build(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=195, corner_radius=0)
        sidebar.grid(row=0, column=0, sticky="nsew")
        sidebar.grid_propagate(False)

        ctk.CTkLabel(sidebar, text="Quem Vai?", font=F_TITLE).pack(
            pady=(32, 4), padx=16
        )
        cfg = _load_config()
        nome_user = cfg.get("utilizador", "")
        self.lbl_user = ctk.CTkLabel(
            sidebar, text=nome_user, font=F_SMALL, text_color="gray50"
        )
        self.lbl_user.pack(pady=(0, 32), padx=16)

        self.btn_sorteio = ctk.CTkButton(
            sidebar, text="Sorteio", font=F_MEDIUM, height=BTN_H,
            anchor="w", fg_color=NAV_INACTIVE,
            hover_color=("gray85", "gray30"),
            command=lambda: self._nav("sorteio")
        )
        self.btn_sorteio.pack(fill="x", padx=12, pady=4)

        self.btn_config = ctk.CTkButton(
            sidebar, text="Configurar", font=F_MEDIUM, height=BTN_H,
            anchor="w", fg_color=NAV_INACTIVE,
            hover_color=("gray85", "gray30"),
            command=lambda: self._nav("config")
        )
        self.btn_config.pack(fill="x", padx=12, pady=4)

        self.frame_sorteio = SorteioFrame(self, self.sorteio)
        self.frame_config  = ConfigFrame(self)
        self.frame_sorteio.grid(row=0, column=1, sticky="nsew")
        self.frame_config.grid(row=0, column=1, sticky="nsew")

    def _setup_primeiro_uso(self):
        d = DialogoPrimeiroUso(self)
        self.wait_window(d)
        if d.nome:
            _save_config({"utilizador": d.nome})
            self.lbl_user.configure(text=d.nome)

    def _nav(self, destino: str):
        self.frame_sorteio.grid_remove()
        self.frame_config.grid_remove()
        if destino == "sorteio":
            self.frame_sorteio.grid()
            self.frame_sorteio.on_show()
            self.btn_sorteio.configure(fg_color=NAV_ACTIVE)
            self.btn_config.configure(fg_color=NAV_INACTIVE)
        else:
            self.frame_config.grid()
            self.frame_config.on_show()
            self.btn_config.configure(fg_color=NAV_ACTIVE)
            self.btn_sorteio.configure(fg_color=NAV_INACTIVE)


# ─── Arranque ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import tkinter as _tk
    _root = _tk.Tk()
    _root.withdraw()
    _setup_treeview_style()
    primeiro_uso = not os.path.isdir(_app_folder())
    _init_db()
    QuemVaiApp(primeiro_uso=primeiro_uso).mainloop()
