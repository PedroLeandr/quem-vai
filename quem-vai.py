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
    return os.path.dirname(
        sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__)
    )


def _db_path() -> str:
    return os.path.join(_app_folder(), "quem-vai.db")


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
            CREATE TABLE IF NOT EXISTS atividades (
                id   INTEGER PRIMARY KEY AUTOINCREMENT,
                nome TEXT NOT NULL UNIQUE
            );
            CREATE TABLE IF NOT EXISTS atividade_equipas (
                atividade_id INTEGER NOT NULL REFERENCES atividades(id) ON DELETE CASCADE,
                equipa_id    INTEGER NOT NULL REFERENCES equipas(id)    ON DELETE CASCADE,
                num_alunos   INTEGER NOT NULL DEFAULT 1,
                PRIMARY KEY (atividade_id, equipa_id)
            );
            CREATE TABLE IF NOT EXISTS participacoes (
                atividade_id INTEGER NOT NULL REFERENCES atividades(id) ON DELETE CASCADE,
                jogador_id   INTEGER NOT NULL REFERENCES jogadores(id)  ON DELETE CASCADE,
                PRIMARY KEY (atividade_id, jogador_id)
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
        rows = c.execute(
            "SELECT e.id, e.nome, e.tipo, COUNT(j.id) AS total"
            " FROM equipas e LEFT JOIN jogadores j ON j.equipa_id=e.id"
            " GROUP BY e.id"
        ).fetchall()
    return sorted(rows, key=lambda e: e["nome"].casefold())


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


# ─── Helpers DB: Atividades ───────────────────────────────────────────────────

def listar_atividades() -> list[dict]:
    with _conn() as c:
        rows = c.execute("SELECT id, nome FROM atividades").fetchall()
    return sorted(rows, key=lambda a: a["nome"].casefold())


def criar_atividade(nome: str) -> None:
    with _conn() as c:
        c.execute("INSERT INTO atividades (nome) VALUES (?)", (nome,))


def apagar_atividade(aid: int) -> None:
    with _conn() as c:
        c.execute("DELETE FROM atividades WHERE id=?", (aid,))


def renomear_atividade(aid: int, nome: str) -> None:
    with _conn() as c:
        c.execute("UPDATE atividades SET nome=? WHERE id=?", (nome, aid))


def get_config_atividade(aid: int) -> list[dict]:
    equipas = listar_equipas()
    with _conn() as c:
        cfg = {row["equipa_id"]: row["num_alunos"] for row in c.execute(
            "SELECT equipa_id, num_alunos FROM atividade_equipas WHERE atividade_id=?", (aid,)
        ).fetchall()}
        participaram = {row["equipa_id"]: row["cnt"] for row in c.execute(
            "SELECT j.equipa_id, COUNT(*) AS cnt"
            " FROM participacoes p JOIN jogadores j ON j.id=p.jogador_id"
            " WHERE p.atividade_id=? GROUP BY j.equipa_id",
            (aid,)
        ).fetchall()}
    for eq in equipas:
        eq["num_alunos"]  = cfg.get(eq["id"], 0)
        eq["participaram"] = participaram.get(eq["id"], 0)
        eq["restantes"]   = eq["total"] - eq["participaram"]
    return equipas


def set_num_alunos(aid: int, equipa_id: int, num: int) -> None:
    with _conn() as c:
        c.execute(
            "INSERT OR REPLACE INTO atividade_equipas (atividade_id, equipa_id, num_alunos)"
            " VALUES (?,?,?)",
            (aid, equipa_id, num)
        )


def sortear_atividade(aid: int, pre_selected: dict[int, list[int]] | None = None) -> list[dict]:
    pre_selected = pre_selected or {}

    with _conn() as c:
        config = {row["equipa_id"]: row for row in c.execute(
            "SELECT ae.equipa_id, ae.num_alunos, e.nome AS equipa_nome, e.tipo"
            " FROM atividade_equipas ae JOIN equipas e ON e.id=ae.equipa_id"
            " WHERE ae.atividade_id=?",
            (aid,)
        ).fetchall()}

    # Incluir equipas com apenas seleção manual
    for eid in pre_selected:
        if eid not in config:
            with _conn() as c:
                row = c.execute(
                    "SELECT id AS equipa_id, nome AS equipa_nome, tipo FROM equipas WHERE id=?", (eid,)
                ).fetchone()
            if row:
                config[eid] = {**row, "num_alunos": 0}

    resultados = []
    for eid, cfg in config.items():
        n = cfg["num_alunos"]
        manual_ids = set(pre_selected.get(eid, []))
        effective_n = max(n, len(manual_ids))
        if effective_n == 0:
            continue

        def _get_disponiveis_nao_manual():
            with _conn() as c:
                return [j for j in c.execute(
                    "SELECT j.id, j.nome, j.numero_interno FROM jogadores j"
                    " WHERE j.equipa_id=?"
                    "   AND j.id NOT IN (SELECT jogador_id FROM participacoes WHERE atividade_id=?)",
                    (eid, aid)
                ).fetchall() if j["id"] not in manual_ids]

        disponiveis = _get_disponiveis_nao_manual()
        needed_random = effective_n - len(manual_ids)

        if needed_random > len(disponiveis):
            with _conn() as c:
                c.execute(
                    "DELETE FROM participacoes WHERE atividade_id=?"
                    " AND jogador_id IN (SELECT id FROM jogadores WHERE equipa_id=?)",
                    (aid, eid)
                )
            disponiveis = _get_disponiveis_nao_manual()

        random_sel = random.sample(disponiveis, min(needed_random, len(disponiveis)))

        manual_players = []
        if manual_ids:
            with _conn() as c:
                for mid in manual_ids:
                    row = c.execute(
                        "SELECT id, nome, numero_interno FROM jogadores WHERE id=?", (mid,)
                    ).fetchone()
                    if row:
                        manual_players.append(row)

        selecionados = manual_players + random_sel
        with _conn() as c:
            for j in selecionados:
                c.execute(
                    "INSERT OR IGNORE INTO participacoes (atividade_id, jogador_id) VALUES (?,?)",
                    (aid, j["id"])
                )
        for j in selecionados:
            resultados.append({
                "nome": j["nome"],
                "numero_interno": j.get("numero_interno"),
                "equipa": cfg["equipa_nome"],
                "tipo": cfg["tipo"],
            })

    return resultados


def reset_atividade(aid: int) -> None:
    with _conn() as c:
        c.execute("DELETE FROM participacoes WHERE atividade_id=?", (aid,))


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
            columns=("num", "nome", "del"),
            show="headings",
            style="Excel.Treeview",
            selectmode="browse",
        )
        self.tree.heading("num",  text="Nº Interno", anchor="w")
        self.tree.heading("nome", text="Nome",        anchor="w")
        self.tree.heading("del",  text="",            anchor="center")
        self.tree.column("num",  width=130, minwidth=80,  stretch=False)
        self.tree.column("nome", width=340, minwidth=120, stretch=True)
        self.tree.column("del",  width=44,  minwidth=44,  stretch=False)

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
        self.tree.tag_configure("del_hover", foreground="#ef4444")

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
                             values=(a["numero_interno"] or "", a["nome"], "🗑"),
                             tags=tags)
        self.tree.insert("", "end", iid=self.NOVO_ID,
                         values=("", "← escreve aqui para adicionar aluno", ""),
                         tags=("novo",))

    # ── Interação ────────────────────────────────────────────────────────────

    def _on_click(self, event):
        item = self.tree.identify_row(event.y)
        col  = self.tree.identify_column(event.x)
        if not item or not col:
            return
        if col == "#3" and item != self.NOVO_ID:
            self._delete_row(int(item))
        else:
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


# ─── Diálogo de Seleção Manual ───────────────────────────────────────────────

class DialogoSelecaoManual(ctk.CTkToplevel):
    def __init__(self, parent, equipa: dict, aid: int, pre_selected: list[int]):
        super().__init__(parent)
        self.title(f"Selecionar — {equipa['nome']}")
        self.geometry("400x480")
        self.resizable(False, True)
        self.selecionados: list[int] = []
        self._aid = aid
        self._vars: dict[int, tk.BooleanVar] = {}

        with _conn() as c:
            jogadores = c.execute(
                "SELECT id, nome, numero_interno FROM jogadores WHERE equipa_id=? ORDER BY nome",
                (equipa["id"],)
            ).fetchall()
            participaram = {r["jogador_id"] for r in c.execute(
                "SELECT jogador_id FROM participacoes WHERE atividade_id=?", (aid,)
            ).fetchall()}

        pre_set = set(pre_selected)

        ctk.CTkLabel(self, text=equipa["nome"], font=F_LARGE).pack(pady=(18, 4))
        ctk.CTkLabel(self, text="Seleciona quem vai neste sorteio", font=F_SMALL,
                     text_color="gray50").pack(pady=(0, 10))

        scroll = ctk.CTkScrollableFrame(self)
        scroll.pack(fill="both", expand=True, padx=16)

        for j in jogadores:
            var = tk.BooleanVar(value=j["id"] in pre_set)
            self._vars[j["id"]] = var
            num = f"  (Nº{j['numero_interno']})" if j.get("numero_interno") else ""
            txt = f"{j['nome']}{num}"
            state = "disabled" if j["id"] in participaram else "normal"
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=2)
            cb = ctk.CTkCheckBox(row, text=txt, variable=var, font=F_SMALL,
                                 state=state, command=self._atualizar_contador)
            cb.pack(side="left", padx=8)
            if j["id"] in participaram:
                ctk.CTkLabel(row, text="já participou", font=("Segoe UI", 11),
                             text_color="gray40").pack(side="right", padx=8)

        self.lbl_contador = ctk.CTkLabel(self, text="", font=F_SMALL, text_color="gray60")
        self.lbl_contador.pack(pady=(8, 4))

        ctk.CTkButton(self, text="Confirmar", font=F_MEDIUM, height=BTN_H,
                      command=self._confirmar).pack(padx=16, pady=(0, 16), fill="x")

        self._atualizar_contador()
        self.after(150, self._on_show)

    def _on_show(self):
        self.lift()
        self.grab_set()

    def _atualizar_contador(self):
        n = sum(1 for v in self._vars.values() if v.get())
        self.lbl_contador.configure(text=f"{n} selecionado{'s' if n != 1 else ''} manualmente")

    def _confirmar(self):
        self.selecionados = [jid for jid, v in self._vars.items() if v.get()]
        self.destroy()


# ─── Diálogo de Atividade ─────────────────────────────────────────────────────

class DialogoAtividade(ctk.CTkToplevel):
    def __init__(self, parent, atividade: dict | None = None):
        super().__init__(parent)
        self.title("Brigada")
        self.geometry("380x175")
        self.resizable(False, False)
        self.resultado = None

        ctk.CTkLabel(self, text="Nome da brigada:", font=F_MEDIUM).pack(pady=(20, 6))
        self.entry = ctk.CTkEntry(self, font=F_MEDIUM, height=BTN_H, width=320)
        self.entry.pack()
        if atividade:
            self.entry.insert(0, atividade["nome"])

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
        self.resultado = nome
        self.destroy()


# ─── Frame: Atividades ────────────────────────────────────────────────────────

class AtividadesFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent, fg_color="transparent")
        self._ativ_sel: dict | None = None
        self._ativ_btns: dict[int, ctk.CTkButton] = {}
        self._entries_num: dict[int, ctk.CTkEntry] = {}
        self._sel_btns: dict[int, ctk.CTkButton] = {}
        self._selecao_manual: dict[int, list[int]] = {}
        self._historico: list[list[dict]] = []
        self.lista_hist = None
        self._build()

    def _build(self):
        header = ctk.CTkFrame(self, fg_color="transparent")
        header.pack(fill="x", padx=24, pady=(24, 12))
        ctk.CTkLabel(header, text="Brigadas", font=F_LARGE).pack(side="left")
        ctk.CTkButton(
            header, text="+ Brigada", font=F_SMALL, height=40, width=140,
            command=self._nova_ativ
        ).pack(side="right")

        paineis = ctk.CTkFrame(self, fg_color="transparent")
        paineis.pack(fill="both", expand=True, padx=24, pady=(0, 24))
        paineis.columnconfigure(0, weight=2)
        paineis.columnconfigure(1, weight=3)
        paineis.rowconfigure(0, weight=1)

        self.lista_ativs = ctk.CTkScrollableFrame(paineis, label_text="Brigadas")
        self.lista_ativs.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.painel_dir = ctk.CTkFrame(paineis)
        self.painel_dir.grid(row=0, column=1, sticky="nsew")

        self._placeholder_ativ()

    def _placeholder_ativ(self):
        for w in self.painel_dir.winfo_children():
            w.destroy()
        self.lista_hist = None
        ctk.CTkLabel(
            self.painel_dir,
            text="Seleciona uma brigada para configurar",
            font=F_MEDIUM, text_color="gray50"
        ).pack(expand=True)

    def on_show(self):
        self._refresh_ativs()

    # ── Lista de atividades ───────────────────────────────────────────────────

    def _refresh_ativs(self):
        self._ativ_btns.clear()
        for w in self.lista_ativs.winfo_children():
            w.destroy()
        for at in listar_atividades():
            self._linha_ativ(at)
        self._atualizar_cor_btns()

    def _linha_ativ(self, at: dict):
        row = ctk.CTkFrame(self.lista_ativs, fg_color="transparent")
        row.pack(fill="x", pady=3)
        btn = ctk.CTkButton(
            row, text=at["nome"],
            font=F_SMALL, anchor="w", height=48,
            command=lambda a=at: self._sel_ativ(a)
        )
        btn.pack(side="left", fill="x", expand=True)
        self._ativ_btns[at["id"]] = btn
        ctk.CTkButton(row, text="✎", width=40, height=48, font=F_SMALL,
                      fg_color="transparent", border_width=1,
                      command=lambda a=at: self._editar_ativ(a)).pack(side="left", padx=2)
        ctk.CTkButton(row, text="✕", width=40, height=48, font=F_SMALL,
                      fg_color="transparent", border_width=1,
                      text_color=("red", "#ef4444"),
                      command=lambda a=at: self._apagar_ativ(a)).pack(side="left")

    def _nova_ativ(self):
        d = DialogoAtividade(self)
        self.wait_window(d)
        if not d.resultado:
            return
        try:
            criar_atividade(d.resultado)
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", f"Já existe uma brigada chamada '{d.resultado}'.")
            return
        self._refresh_ativs()

    def _editar_ativ(self, at: dict):
        d = DialogoAtividade(self, at)
        self.wait_window(d)
        if not d.resultado:
            return
        try:
            renomear_atividade(at["id"], d.resultado)
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", f"Já existe uma brigada chamada '{d.resultado}'.")
            return
        self._refresh_ativs()
        if self._ativ_sel and self._ativ_sel["id"] == at["id"]:
            self._ativ_sel = {**self._ativ_sel, "nome": d.resultado}
            self._refresh_painel_ativ()

    def _apagar_ativ(self, at: dict):
        if not messagebox.askyesno("Confirmar", f"Apagar '{at['nome']}'?"):
            return
        apagar_atividade(at["id"])
        if self._ativ_sel and self._ativ_sel["id"] == at["id"]:
            self._ativ_sel = None
            self._historico.clear()
            self._placeholder_ativ()
        self._refresh_ativs()

    def _atualizar_cor_btns(self):
        sel_id = self._ativ_sel["id"] if self._ativ_sel else None
        for aid, btn in self._ativ_btns.items():
            if aid == sel_id:
                btn.configure(fg_color=NAV_ACTIVE, hover_color=("#1e40af", "#1e40af"))
            else:
                btn.configure(fg_color=("#3b3b3b", "#3b3b3b"), hover_color=("gray85", "gray30"))

    def _sel_ativ(self, at: dict):
        self._ativ_sel = at
        self._historico.clear()
        self._selecao_manual.clear()
        self._atualizar_cor_btns()
        self._refresh_painel_ativ()

    # ── Painel direito ────────────────────────────────────────────────────────

    def _refresh_painel_ativ(self):
        for w in self.painel_dir.winfo_children():
            w.destroy()
        self._entries_num.clear()
        self._sel_btns.clear()
        self.lista_hist = None

        at = self._ativ_sel
        if at is None:
            self._placeholder_ativ()
            return

        equipas = get_config_atividade(at["id"])

        # Título
        topo = ctk.CTkFrame(self.painel_dir, fg_color="transparent")
        topo.pack(fill="x", padx=16, pady=(16, 8))
        ctk.CTkLabel(topo, text=at["nome"], font=F_LARGE).pack(side="left")

        # Cards de equipas num frame scrollável
        scroll_eq = ctk.CTkScrollableFrame(self.painel_dir, height=min(240, max(80, len(equipas) * 58)))
        scroll_eq.pack(fill="x", padx=16, pady=(0, 8))

        for eq in equipas:
            card = ctk.CTkFrame(scroll_eq, corner_radius=8)
            card.pack(fill="x", pady=3)

            # Lado esquerdo: nome + contador
            left = ctk.CTkFrame(card, fg_color="transparent")
            left.pack(side="left", fill="x", expand=True, padx=12, pady=8)
            ctk.CTkLabel(left, text=f"{eq['nome']}  ({eq['tipo']})",
                         font=F_SMALL, anchor="w").pack(anchor="w")
            ctk.CTkLabel(left,
                         text=f"Faltam {eq['restantes']} de {eq['total']}",
                         font=("Segoe UI", 12), text_color="gray50", anchor="w").pack(anchor="w")

            # Lado direito: botão escolher + entrada nº alunos
            right = ctk.CTkFrame(card, fg_color="transparent")
            right.pack(side="right", padx=12)

            n_manual = len(self._selecao_manual.get(eq["id"], []))
            sel_txt = f"✓ {n_manual}" if n_manual else "Escolher"
            sel_btn = ctk.CTkButton(
                right, text=sel_txt, width=90, height=36, font=F_SMALL,
                fg_color=NAV_ACTIVE if n_manual else "transparent",
                border_width=1,
                hover_color=("#1e40af", "#1e40af") if n_manual else ("gray85", "gray30"),
                command=lambda e=eq: self._abrir_selecao(e)
            )
            sel_btn.pack(side="left", padx=(0, 8))
            self._sel_btns[eq["id"]] = sel_btn

            ctk.CTkLabel(right, text="Nº alunos:", font=F_SMALL).pack(side="left", padx=(0, 6))
            entry = ctk.CTkEntry(right, width=64, height=36, font=F_SMALL, justify="center")
            entry.insert(0, str(eq["num_alunos"]))
            entry.pack(side="left")
            self._entries_num[eq["id"]] = entry

            entry.bind("<FocusOut>", lambda e, eid=eq["id"]: self._save_num(eid))
            entry.bind("<Return>",   lambda e, eid=eq["id"]: self._save_num(eid))

        # Botões
        ctk.CTkButton(
            self.painel_dir, text="SORTEAR", font=("Segoe UI", 24, "bold"),
            height=72, corner_radius=12, command=self._sortear
        ).pack(fill="x", padx=16, pady=(8, 4))

        ctk.CTkButton(
            self.painel_dir, text="Reiniciar sessão", font=F_SMALL, height=38,
            fg_color="transparent", border_width=1,
            hover_color=("gray85", "gray30"),
            command=self._reiniciar
        ).pack(padx=16, pady=(0, 8))

        ctk.CTkLabel(self.painel_dir, text="Histórico desta sessão",
                     font=F_MEDIUM).pack(anchor="w", padx=16)

        self.lista_hist = ctk.CTkScrollableFrame(self.painel_dir)
        self.lista_hist.pack(fill="both", expand=True, padx=16, pady=(4, 16))

        self._refresh_historico()

    def _abrir_selecao(self, eq: dict):
        if self._ativ_sel is None:
            return
        pre = self._selecao_manual.get(eq["id"], [])
        d = DialogoSelecaoManual(self, eq, self._ativ_sel["id"], pre)
        self.wait_window(d)
        self._selecao_manual[eq["id"]] = d.selecionados

        # Actualiza botão
        btn = self._sel_btns.get(eq["id"])
        if btn:
            n = len(d.selecionados)
            btn.configure(
                text=f"✓ {n}" if n else "Escolher",
                fg_color=NAV_ACTIVE if n else "transparent",
                hover_color=("#1e40af", "#1e40af") if n else ("gray85", "gray30"),
            )

        # Se manual > num_alunos configurado, aumenta automaticamente
        entry = self._entries_num.get(eq["id"])
        if entry:
            try:
                current = int(entry.get())
            except ValueError:
                current = 0
            n = len(d.selecionados)
            if n > current:
                entry.delete(0, "end")
                entry.insert(0, str(n))
                set_num_alunos(self._ativ_sel["id"], eq["id"], n)

    def _save_num(self, equipa_id: int):
        entry = self._entries_num.get(equipa_id)
        if entry is None or self._ativ_sel is None:
            return
        try:
            val = max(0, int(entry.get()))
        except ValueError:
            val = 0
        entry.delete(0, "end")
        entry.insert(0, str(val))
        set_num_alunos(self._ativ_sel["id"], equipa_id, val)

    def _sortear(self):
        if self._ativ_sel is None:
            return
        for eid in list(self._entries_num):
            self._save_num(eid)
        resultados = sortear_atividade(self._ativ_sel["id"], self._selecao_manual)
        self._selecao_manual.clear()
        if not resultados:
            messagebox.showinfo(
                "Sem alunos",
                "Nenhuma equipa tem alunos configurados.\n"
                "Define o número de alunos por equipa antes de sortear a brigada."
            )
            return
        self._historico.append(resultados)
        self._refresh_painel_ativ()

    def _reiniciar(self):
        if self._ativ_sel is None:
            return
        reset_atividade(self._ativ_sel["id"])
        self._historico.clear()
        self._refresh_painel_ativ()

    def _refresh_historico(self):
        if self.lista_hist is None:
            return
        for w in self.lista_hist.winfo_children():
            w.destroy()
        for ronda in reversed(self._historico):
            por_equipa: dict[str, list[str]] = {}
            for r in ronda:
                num = f" (Nº{r['numero_interno']})" if r.get("numero_interno") else ""
                por_equipa.setdefault(r["equipa"], []).append(f"{r['nome']}{num}")

            frame_ronda = ctk.CTkFrame(self.lista_hist, fg_color="#1e1e2e", corner_radius=8)
            frame_ronda.pack(fill="x", padx=2, pady=4)
            for equipa, nomes in por_equipa.items():
                ctk.CTkLabel(
                    frame_ronda,
                    text=f"{equipa}: {', '.join(nomes)}",
                    font=F_SMALL, anchor="w", wraplength=420
                ).pack(fill="x", padx=10, pady=3)


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
        self.after(10, lambda: self.state("zoomed"))
        self._build()
        self._nav("atividades")
        self.protocol("WM_DELETE_WINDOW", self._on_close)
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

        self.btn_atividades = ctk.CTkButton(
            sidebar, text="Brigadas", font=F_MEDIUM, height=BTN_H,
            anchor="w", fg_color=NAV_INACTIVE,
            hover_color=("gray85", "gray30"),
            command=lambda: self._nav("atividades")
        )
        self.btn_atividades.pack(fill="x", padx=12, pady=4)

        self.btn_config = ctk.CTkButton(
            sidebar, text="Configurar", font=F_MEDIUM, height=BTN_H,
            anchor="w", fg_color=NAV_INACTIVE,
            hover_color=("gray85", "gray30"),
            command=lambda: self._nav("config")
        )
        self.btn_config.pack(fill="x", padx=12, pady=4)

        self.frame_atividades = AtividadesFrame(self)
        self.frame_config     = ConfigFrame(self)
        self.frame_atividades.grid(row=0, column=1, sticky="nsew")
        self.frame_config.grid(row=0, column=1, sticky="nsew")

    def _on_close(self):
        self.quit()
        self.destroy()

    def _setup_primeiro_uso(self):
        d = DialogoPrimeiroUso(self)
        self.wait_window(d)
        if d.nome:
            _save_config({"utilizador": d.nome})
            self.lbl_user.configure(text=d.nome)

    def _nav(self, destino: str):
        self.frame_atividades.grid_remove()
        self.frame_config.grid_remove()
        if destino == "atividades":
            self.frame_atividades.grid()
            self.frame_atividades.on_show()
            self.btn_atividades.configure(fg_color=NAV_ACTIVE)
            self.btn_config.configure(fg_color=NAV_INACTIVE)
        else:
            self.frame_config.grid()
            self.frame_config.on_show()
            self.btn_config.configure(fg_color=NAV_ACTIVE)
            self.btn_atividades.configure(fg_color=NAV_INACTIVE)


# ─── Arranque ─────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import tkinter as _tk
    _root = _tk.Tk()
    _root.withdraw()
    _setup_treeview_style()
    primeiro_uso = not os.path.isdir(_app_folder())
    _init_db()
    QuemVaiApp(primeiro_uso=primeiro_uso).mainloop()
