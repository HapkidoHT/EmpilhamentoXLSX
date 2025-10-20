import os
import threading
import queue
import traceback
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

APP_TITLE = "Empilhador de Planilhas Excel"
DEFAULT_SHEET = "Plan1"
VALID_EXTS = {".xlsx", ".xlsm"}

# ---------------------------
# Utilidades
# ---------------------------
def is_excel_file(path: str) -> bool:
    name = os.path.basename(path)
    if name.startswith("~$"):  # ignora arquivos temporários do Office
        return False
    _, ext = os.path.splitext(name)
    return ext.lower() in VALID_EXTS

def list_excel_files(root_dir: str, recursive: bool) -> list[str]:
    files = []
    if recursive:
        for r, _, fns in os.walk(root_dir):
            for fn in fns:
                full = os.path.join(r, fn)
                if is_excel_file(full):
                    files.append(full)
    else:
        for fn in os.listdir(root_dir):
            full = os.path.join(root_dir, fn)
            if os.path.isfile(full) and is_excel_file(full):
                files.append(full)
    files.sort()
    return files

def safe_concat(dataframes: list[pd.DataFrame]) -> pd.DataFrame:
    if not dataframes:
        return pd.DataFrame()
    # União de colunas: reindex em todas as colunas encontradas, mantendo ordem estável
    all_cols = []
    seen = set()
    for df in dataframes:
        for c in df.columns:
            if c not in seen:
                seen.add(c)
                all_cols.append(c)
    reindexed = [df.reindex(columns=all_cols) for df in dataframes]
    return pd.concat(reindexed, ignore_index=True)

# ---------------------------
# Worker de combinação
# ---------------------------
class CombinerWorker(threading.Thread):
    def __init__(self, directory, sheet_name, recursive, add_filename_col, out_path, progress_cb, log_q, stop_flag):
        super().__init__(daemon=True)
        self.directory = directory
        self.sheet_name = sheet_name
        self.recursive = recursive
        self.add_filename_col = add_filename_col
        self.out_path = out_path
        self.progress_cb = progress_cb
        self.log_q = log_q
        self.stop_flag = stop_flag

    def log(self, msg: str):
        self.log_q.put(msg)

    def run(self):
        try:
            files = list_excel_files(self.directory, self.recursive)
            total = len(files)
            if total == 0:
                self.log("Nenhum arquivo .xlsx/.xlsm encontrado.")
                self.progress_cb(0, "Pronto (0 arquivos).")
                return

            self.log(f"Arquivos candidatos encontrados: {total}")
            data_frames = []
            processed = 0
            with_sheet = 0
            empty_list = []

            for idx, path in enumerate(files, start=1):
                if self.stop_flag.is_set():
                    self.log("Processo cancelado pelo usuário.")
                    self.progress_cb(int(idx / total * 100), "Cancelado.")
                    return

                base = os.path.basename(path)
                try:
                    # Checa abas rapidamente
                    with pd.ExcelFile(path) as xf:
                        if self.sheet_name not in xf.sheet_names:
                            self.log(f"[PULADO] '{base}' não tem a aba '{self.sheet_name}'.")
                            self.progress_cb(int(idx / total * 100), f"Pulado: {base}")
                            continue

                        with_sheet += 1
                        df = pd.read_excel(xf, sheet_name=self.sheet_name)  # header padrão
                except Exception as e:
                    self.log(f"[ERRO] Falha ao abrir/ler '{base}': {e}")
                    self.progress_cb(int(idx / total * 100), f"Erro: {base}")
                    continue

                if df is None or df.empty:
                    empty_list.append(base)
                else:
                    if self.add_filename_col:
                        df.insert(0, "Arquivo_Origem", base)
                    data_frames.append(df)

                processed += 1
                self.progress_cb(int(idx / total * 100), f"Lido: {base}")

            self.log(f"Arquivos lidos: {processed} | Possuíam a aba '{self.sheet_name}': {with_sheet}")
            if empty_list:
                self.log(f"Atenção: {len(empty_list)} arquivos com aba '{self.sheet_name}' vazia: {', '.join(empty_list)}")

            if not data_frames:
                self.log("Nenhum dado para combinar (todas as abas estavam vazias ou inexistentes).")
                self.progress_cb(100, "Pronto (sem dados).")
                return

            combined = safe_concat(data_frames)

            # Salvar
            try:
                # Garante pasta de destino
                os.makedirs(os.path.dirname(self.out_path), exist_ok=True)
                combined.to_excel(self.out_path, index=False)
            except Exception as e:
                self.log(f"[ERRO] Falha ao salvar o arquivo final: {e}")
                self.progress_cb(100, "Erro ao salvar.")
                return

            self.log(f"Sucesso! Linhas combinadas: {len(combined)}")
            self.log(f"Arquivo salvo em: {self.out_path}")
            self.progress_cb(100, "Concluído.")
        except Exception as e:
            self.log("[ERRO FATAL] Ocorreu um erro inesperado.")
            self.log(str(e))
            self.log(traceback.format_exc())
            self.progress_cb(100, "Erro.")

# ---------------------------
# Aplicação (UI)
# ---------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("760x560")
        self.minsize(720, 520)

        # Estilo TTK
        try:
            self.style = ttk.Style(self)
            # Usa tema disponível
            if "clam" in self.style.theme_names():
                self.style.theme_use("clam")
            self.style.configure("TFrame", padding=10)
            self.style.configure("TButton", padding=6)
            self.style.configure("TEntry", padding=4)
            self.style.configure("TLabel", padding=4)
            self.style.configure("Horizontal.TProgressbar", thickness=16)
        except Exception:
            pass

        self.directory = tk.StringVar()
        self.sheet_name = tk.StringVar(value=DEFAULT_SHEET)
        self.recursive = tk.BooleanVar(value=False)
        self.add_filename_col = tk.BooleanVar(value=True)
        self.output_path = tk.StringVar(value="")

        self.log_q = queue.Queue()
        self.stop_flag = threading.Event()
        self.worker = None

        self._build_ui()
        self._after_poll_log()

    def _build_ui(self):
        container = ttk.Frame(self)
        container.pack(fill="both", expand=True)

        # Seleção de pasta
        frm_dir = ttk.LabelFrame(container, text="Origem")
        frm_dir.pack(fill="x", padx=4, pady=6)
        row1 = ttk.Frame(frm_dir)
        row1.pack(fill="x", pady=4)
        ttk.Label(row1, text="Pasta com as planilhas:").pack(side="left")
        ent_dir = ttk.Entry(row1, textvariable=self.directory)
        ent_dir.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row1, text="Selecionar pasta…", command=self.on_pick_folder).pack(side="left")

        # Parâmetros
        frm_params = ttk.LabelFrame(container, text="Parâmetros")
        frm_params.pack(fill="x", padx=4, pady=6)

        row2 = ttk.Frame(frm_params)
        row2.pack(fill="x", pady=4)
        ttk.Label(row2, text="Nome da aba a empilhar:").pack(side="left")
        ttk.Entry(row2, width=24, textvariable=self.sheet_name).pack(side="left", padx=6)
        ttk.Checkbutton(row2, text="Varrer subpastas", variable=self.recursive).pack(side="left", padx=12)
        ttk.Checkbutton(row2, text="Adicionar coluna 'Arquivo_Origem'", variable=self.add_filename_col).pack(side="left", padx=12)

        # Saída
        frm_out = ttk.LabelFrame(container, text="Saída")
        frm_out.pack(fill="x", padx=4, pady=6)
        row3 = ttk.Frame(frm_out)
        row3.pack(fill="x", pady=4)
        ttk.Label(row3, text="Arquivo final:").pack(side="left")
        ent_out = ttk.Entry(row3, textvariable=self.output_path)
        ent_out.pack(side="left", fill="x", expand=True, padx=6)
        ttk.Button(row3, text="Salvar como…", command=self.on_pick_save).pack(side="left")

        # Ações
        frm_actions = ttk.Frame(container)
        frm_actions.pack(fill="x", padx=4, pady=6)
        self.btn_start = ttk.Button(frm_actions, text="Empilhar planilhas", command=self.on_start)
        self.btn_start.pack(side="left")
        self.btn_stop = ttk.Button(frm_actions, text="Cancelar", command=self.on_stop, state="disabled")
        self.btn_stop.pack(side="left", padx=8)

        # Progresso
        frm_prog = ttk.Frame(container)
        frm_prog.pack(fill="x", padx=4, pady=6)
        self.progress = ttk.Progressbar(frm_prog, mode="determinate", maximum=100)
        self.progress.pack(fill="x")
        self.lbl_status = ttk.Label(frm_prog, text="Pronto.")
        self.lbl_status.pack(anchor="w", pady=2)

        # Log
        frm_log = ttk.LabelFrame(container, text="Log")
        frm_log.pack(fill="both", expand=True, padx=4, pady=6)
        self.txt_log = tk.Text(frm_log, height=12, wrap="word", state="disabled")
        self.txt_log.pack(fill="both", expand=True)
        # Scrollbar
        sb = ttk.Scrollbar(self.txt_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log["yscrollcommand"] = sb.set
        sb.pack(side="right", fill="y")

        # Rodapé
        ttk.Label(container, text="Formatos aceitos: .xlsx e .xlsm | Usa união de colunas automaticamente.",
                  foreground="#666").pack(anchor="w", padx=6, pady=(0,6))

    # ---------------------------
    # Handlers
    # ---------------------------
    def on_pick_folder(self):
        d = filedialog.askdirectory(title="Selecione a pasta com as planilhas")
        if d:
            self.directory.set(d)
            # Sugere um caminho de saída padrão
            default_out = os.path.join(d, "planilha_combinada.xlsx")
            if not self.output_path.get():
                self.output_path.set(default_out)

    def on_pick_save(self):
        initial = self.output_path.get() or "planilha_combinada.xlsx"
        f = filedialog.asksaveasfilename(
            title="Salvar arquivo combinado",
            defaultextension=".xlsx",
            initialfile=os.path.basename(initial),
            filetypes=[("Excel", "*.xlsx")]
        )
        if f:
            self.output_path.set(f)

    def on_start(self):
        directory = self.directory.get().strip()
        sheet = self.sheet_name.get().strip()
        out = self.output_path.get().strip()

        if not directory or not os.path.isdir(directory):
            messagebox.showwarning("Atenção", "Selecione uma pasta válida.")
            return
        if not sheet:
            messagebox.showwarning("Atenção", "Informe o nome da aba a empilhar.")
            return
        if not out:
            messagebox.showwarning("Atenção", "Escolha onde salvar o arquivo final.")
            return

        # Reset UI
        self.clear_log()
        self.set_status(0, "Iniciando…")
        self.disable_inputs(True)
        self.stop_flag.clear()

        # Start worker thread
        self.worker = CombinerWorker(
            directory=directory,
            sheet_name=sheet,
            recursive=self.recursive.get(),
            add_filename_col=self.add_filename_col.get(),
            out_path=out,
            progress_cb=self.update_progress_threadsafe,
            log_q=self.log_q,
            stop_flag=self.stop_flag
        )
        self.worker.start()

    def on_stop(self):
        if self.worker and self.worker.is_alive():
            self.stop_flag.set()
            self.append_log("Solicitando cancelamento…")

    # ---------------------------
    # UI helpers
    # ---------------------------
    def disable_inputs(self, running: bool):
        state = "disabled" if running else "normal"
        for child in self.winfo_children():
            # controla somente botões principais e entradas
            pass
        self.btn_start.configure(state="disabled" if running else "normal")
        self.btn_stop.configure(state="normal" if running else "disabled")

    def set_status(self, pct: int, msg: str):
        self.progress["value"] = pct
        self.lbl_status.configure(text=msg)
        self.update_idletasks()
        if pct >= 100:
            self.disable_inputs(False)

    def update_progress_threadsafe(self, pct: int, msg: str):
        # chamado de thread; usa after para segurança
        self.after(0, lambda: self.set_status(pct, msg))

    def append_log(self, text: str):
        self.txt_log.configure(state="normal")
        self.txt_log.insert("end", text + "\n")
        self.txt_log.see("end")
        self.txt_log.configure(state="disabled")

    def clear_log(self):
        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

    def _after_poll_log(self):
        # Puxa mensagens do worker e joga no log
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.append_log(msg)
        except queue.Empty:
            pass
        self.after(100, self._after_poll_log)

# ---------------------------
# Main
# ---------------------------
if __name__ == "__main__":
    app = App()
    app.mainloop()
