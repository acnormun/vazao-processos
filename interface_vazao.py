#!/usr/bin/env python3
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import cruzar_vazao


class VazaoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cruzamento de Vazao de Processos")
        self.geometry("760x340")
        self.minsize(680, 300)

        self.distribuicoes_var = tk.StringVar()
        self.saidas_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "vazao-processos.xlsx"))
        self.status_var = tk.StringVar(value="Selecione as duas planilhas para gerar o arquivo.")

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)

        ttk.Label(root, text="Distribuicoes / entrada (.csv)").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Entry(root, textvariable=self.distribuicoes_var).grid(
            row=1, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(root, text="Selecionar", command=self.select_distribuicoes).grid(
            row=1, column=2, sticky="ew"
        )

        ttk.Label(root, text="Saidas (.xlsx)").grid(row=2, column=0, sticky="w", pady=(16, 6))
        ttk.Entry(root, textvariable=self.saidas_var).grid(
            row=3, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(root, text="Selecionar", command=self.select_saidas).grid(
            row=3, column=2, sticky="ew"
        )

        ttk.Label(root, text="Arquivo de resultado (.xlsx ou .csv)").grid(
            row=4, column=0, sticky="w", pady=(16, 6)
        )
        ttk.Entry(root, textvariable=self.output_var).grid(
            row=5, column=0, columnspan=2, sticky="ew", padx=(0, 8)
        )
        ttk.Button(root, text="Salvar como", command=self.select_output).grid(
            row=5, column=2, sticky="ew"
        )

        self.generate_button = ttk.Button(root, text="Gerar planilha", command=self.generate)
        self.generate_button.grid(row=6, column=0, sticky="w", pady=(22, 0))

        ttk.Label(root, textvariable=self.status_var, wraplength=700).grid(
            row=7, column=0, columnspan=3, sticky="ew", pady=(18, 0)
        )

    def select_distribuicoes(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de distribuicoes",
            filetypes=[("CSV", "*.csv"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.distribuicoes_var.set(path)
            self._suggest_output_path(path)

    def select_saidas(self):
        path = filedialog.askopenfilename(
            title="Selecione a planilha de saidas",
            filetypes=[("Excel", "*.xlsx"), ("Todos os arquivos", "*.*")],
        )
        if path:
            self.saidas_var.set(path)

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="Salvar resultado como",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=Path(self.output_var.get()).name,
        )
        if path:
            self.output_var.set(path)

    def _suggest_output_path(self, distribuicoes_path):
        current = Path(self.output_var.get())
        if current.name != "vazao-processos.xlsx":
            return
        folder = Path(distribuicoes_path).parent
        self.output_var.set(str(folder / "vazao-processos.xlsx"))

    def generate(self):
        distribuicoes = Path(self.distribuicoes_var.get().strip())
        saidas = Path(self.saidas_var.get().strip())
        output = Path(self.output_var.get().strip())

        if not distribuicoes.is_file():
            messagebox.showerror("Arquivo invalido", "Selecione o CSV de distribuicoes.")
            return
        if not saidas.is_file():
            messagebox.showerror("Arquivo invalido", "Selecione o XLSX de saidas.")
            return
        if not output.name:
            messagebox.showerror("Saida invalida", "Informe onde salvar o resultado.")
            return

        self.generate_button.configure(state="disabled")
        self.status_var.set("Gerando planilha...")
        thread = threading.Thread(
            target=self._generate_worker,
            args=(distribuicoes, saidas, output),
            daemon=True,
        )
        thread.start()

    def _generate_worker(self, distribuicoes_path, saidas_path, output_path):
        try:
            distribuicoes = cruzar_vazao.read_distribuicoes(distribuicoes_path)
            saidas = cruzar_vazao.read_saidas(saidas_path)
            resultado = cruzar_vazao.cruzar(distribuicoes, saidas)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            cruzar_vazao.write_output(output_path, resultado)

            processos_entrada = set(distribuicoes)
            processos_saida = {saida["processo"] for saida in saidas}
            message = (
                f"Arquivo gerado: {output_path}\n"
                f"Entradas lidas: {cruzar_vazao.count_distribuicoes(distribuicoes)} | "
                f"Processos com entrada: {len(distribuicoes)} | "
                f"Saidas lidas: {len(saidas)} | "
                f"Processos cruzados: {len(processos_entrada & processos_saida)} | "
                f"Entradas ainda sem saida: {cruzar_vazao.count_pendentes(resultado)} | "
                f"Saidas sem entrada: {cruzar_vazao.count_saidas_sem_entrada(resultado)} | "
                f"Linhas geradas: {len(resultado)}"
            )
            self.after(0, self._finish_success, message)
        except Exception as exc:
            self.after(0, self._finish_error, str(exc))

    def _finish_success(self, message):
        self.generate_button.configure(state="normal")
        self.status_var.set(message)
        messagebox.showinfo("Concluido", message)

    def _finish_error(self, message):
        self.generate_button.configure(state="normal")
        self.status_var.set(f"Erro: {message}")
        messagebox.showerror("Erro ao gerar", message)


if __name__ == "__main__":
    app = VazaoApp()
    app.mainloop()
