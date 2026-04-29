#!/usr/bin/env python3
import threading
import tkinter as tk
import datetime as dt
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import cruzar_vazao


class VazaoApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cruzamento de Vazao de Processos")
        self.geometry("1040x760")
        self.minsize(880, 620)

        self.distribuicoes_var = tk.StringVar()
        self.saidas_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "vazao-processos.xlsx"))
        self.status_var = tk.StringVar(value="Selecione as duas planilhas para gerar o arquivo.")
        self.tarefa_var = tk.StringVar(value="Todas")
        self.periodo_var = tk.StringVar(value="Todos")
        self.inicio_var = tk.StringVar()
        self.fim_var = tk.StringVar()
        self.resultado = []

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(8, weight=1)

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

        ttk.Label(root, textvariable=self.status_var, wraplength=960).grid(
            row=7, column=0, columnspan=3, sticky="ew", pady=(18, 10)
        )

        analytics = ttk.LabelFrame(root, text="Vazao por assessor", padding=12)
        analytics.grid(row=8, column=0, columnspan=3, sticky="nsew")
        analytics.columnconfigure(0, weight=1)
        analytics.rowconfigure(2, weight=1)

        filters = ttk.Frame(analytics)
        filters.grid(row=0, column=0, sticky="ew")
        filters.columnconfigure(1, weight=1)

        ttk.Label(filters, text="Tarefa").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.tarefa_combo = ttk.Combobox(
            filters,
            textvariable=self.tarefa_var,
            state="readonly",
            values=["Todas"],
            width=38,
        )
        self.tarefa_combo.grid(row=0, column=1, sticky="ew", padx=(0, 14))

        ttk.Label(filters, text="Periodo").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self.periodo_combo = ttk.Combobox(
            filters,
            textvariable=self.periodo_var,
            state="readonly",
            values=[
                "Todos",
                "Este mes",
                "Ultimos 30 dias",
                "Este ano",
                "Ultimos 12 meses",
                "Personalizado",
            ],
            width=18,
        )
        self.periodo_combo.grid(row=0, column=3, sticky="w", padx=(0, 14))

        ttk.Label(filters, text="Inicio").grid(row=0, column=4, sticky="w", padx=(0, 6))
        ttk.Entry(filters, textvariable=self.inicio_var, width=12).grid(
            row=0, column=5, sticky="w", padx=(0, 10)
        )
        ttk.Label(filters, text="Fim").grid(row=0, column=6, sticky="w", padx=(0, 6))
        ttk.Entry(filters, textvariable=self.fim_var, width=12).grid(
            row=0, column=7, sticky="w", padx=(0, 10)
        )
        ttk.Button(filters, text="Atualizar", command=self.update_analytics).grid(
            row=0, column=8, sticky="e"
        )
        ttk.Button(filters, text="Exportar PDF", command=self.export_pdf).grid(
            row=0, column=9, sticky="e", padx=(8, 0)
        )

        self.summary_var = tk.StringVar(value="Gere a planilha para visualizar os indicadores.")
        ttk.Label(analytics, textvariable=self.summary_var).grid(
            row=1, column=0, sticky="ew", pady=(10, 8)
        )

        content = ttk.PanedWindow(analytics, orient="horizontal")
        content.grid(row=2, column=0, sticky="nsew")

        table_frame = ttk.Frame(content)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        columns = (
            "assessor",
            "saidas",
            "media",
            "mediana",
            "menor",
            "maior",
            "pendentes",
            "sem_entrada",
        )
        self.analytics_table = ttk.Treeview(
            table_frame,
            columns=columns,
            show="headings",
            height=12,
        )
        headings = {
            "assessor": "Assessor",
            "saidas": "Saidas",
            "media": "Media",
            "mediana": "Mediana",
            "menor": "Min",
            "maior": "Max",
            "pendentes": "Sem saida",
            "sem_entrada": "Sem entrada",
        }
        widths = {
            "assessor": 260,
            "saidas": 70,
            "media": 70,
            "mediana": 80,
            "menor": 60,
            "maior": 60,
            "pendentes": 80,
            "sem_entrada": 90,
        }
        for column in columns:
            self.analytics_table.heading(column, text=headings[column])
            self.analytics_table.column(column, width=widths[column], anchor="w")
        self.analytics_table.grid(row=0, column=0, sticky="nsew")
        table_scroll = ttk.Scrollbar(
            table_frame, orient="vertical", command=self.analytics_table.yview
        )
        table_scroll.grid(row=0, column=1, sticky="ns")
        self.analytics_table.configure(yscrollcommand=table_scroll.set)

        chart_frame = ttk.Frame(content)
        chart_frame.columnconfigure(0, weight=1)
        chart_frame.rowconfigure(0, weight=1)
        self.chart = tk.Canvas(chart_frame, background="white", highlightthickness=1)
        self.chart.grid(row=0, column=0, sticky="nsew")
        chart_scroll = ttk.Scrollbar(chart_frame, orient="vertical", command=self.chart.yview)
        chart_scroll.grid(row=0, column=1, sticky="ns")
        self.chart.configure(yscrollcommand=chart_scroll.set)

        content.add(table_frame, weight=3)
        content.add(chart_frame, weight=2)

        self.tarefa_combo.bind("<<ComboboxSelected>>", lambda _event: self.update_analytics())
        self.periodo_combo.bind("<<ComboboxSelected>>", lambda _event: self.update_analytics())
        self.chart.bind("<Configure>", lambda _event: self.update_chart())

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
            self.resultado = resultado
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
        self.refresh_task_filter()
        self.update_analytics()
        messagebox.showinfo("Concluido", message)

    def _finish_error(self, message):
        self.generate_button.configure(state="normal")
        self.status_var.set(f"Erro: {message}")
        messagebox.showerror("Erro ao gerar", message)

    def refresh_task_filter(self):
        tarefas = ["Todas"] + cruzar_vazao.tarefas_disponiveis(self.resultado)
        self.tarefa_combo.configure(values=tarefas)
        if self.tarefa_var.get() not in tarefas:
            self.tarefa_var.set("Todas")

    def parse_filter_date(self, value):
        value = value.strip()
        if not value:
            return None
        try:
            return cruzar_vazao.parse_date(value)
        except ValueError as exc:
            raise ValueError(f"Data de filtro invalida: {value}. Use dd/mm/aaaa.") from exc

    def update_analytics(self):
        if not self.resultado:
            return

        try:
            inicio = self.parse_filter_date(self.inicio_var.get())
            fim = self.parse_filter_date(self.fim_var.get())
        except ValueError as exc:
            messagebox.showerror("Filtro invalido", str(exc))
            return

        tarefa = None if self.tarefa_var.get() == "Todas" else self.tarefa_var.get()
        resumo = cruzar_vazao.resumo_por_assessor(
            self.resultado,
            tarefa=tarefa,
            periodo=self.periodo_var.get(),
            inicio=inicio,
            fim=fim,
        )
        self.current_summary = resumo

        for item_id in self.analytics_table.get_children():
            self.analytics_table.delete(item_id)

        for item in resumo:
            self.analytics_table.insert(
                "",
                "end",
                values=(
                    item["Assessor"],
                    item["Saidas"],
                    item["Vazao media"],
                    item["Vazao mediana"],
                    item["Menor vazao"],
                    item["Maior vazao"],
                    item["Entradas sem saida"],
                    item["Saidas sem entrada"],
                ),
            )

        total_saidas = sum(item["Saidas"] for item in resumo)
        total_pendentes = sum(item["Entradas sem saida"] for item in resumo)
        total_sem_entrada = sum(item["Saidas sem entrada"] for item in resumo)
        self.summary_var.set(
            f"Assessores: {len(resumo)} | Saidas no periodo: {total_saidas} | "
            f"Entradas sem saida: {total_pendentes} | Saidas sem entrada: {total_sem_entrada}"
        )
        self.update_chart()

    def current_filter_text(self):
        parts = [f"Tarefa: {self.tarefa_var.get()}", f"Periodo: {self.periodo_var.get()}"]
        if self.inicio_var.get().strip():
            parts.append(f"Inicio: {self.inicio_var.get().strip()}")
        if self.fim_var.get().strip():
            parts.append(f"Fim: {self.fim_var.get().strip()}")
        return " | ".join(parts)

    def export_pdf(self):
        if not self.resultado:
            messagebox.showerror("Sem dados", "Gere a planilha antes de exportar o PDF.")
            return

        self.update_analytics()
        resumo = getattr(self, "current_summary", [])
        if not resumo:
            messagebox.showerror("Sem dados", "Nao ha dados para os filtros atuais.")
            return

        output = filedialog.asksaveasfilename(
            title="Salvar relatorio em PDF",
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="relatorio-vazao-assessores.pdf",
        )
        if not output:
            return

        try:
            cruzar_vazao.write_pdf_report(
                output,
                resumo,
                filtros=self.current_filter_text(),
            )
        except Exception as exc:
            messagebox.showerror("Erro ao exportar", str(exc))
            return

        messagebox.showinfo("PDF gerado", f"Arquivo gerado: {output}")

    def update_chart(self):
        self.chart.delete("all")
        resumo = getattr(self, "current_summary", [])
        if not resumo:
            self.chart.create_text(20, 20, anchor="nw", text="Sem dados para os filtros atuais.")
            self.chart.configure(scrollregion=(0, 0, 400, 80))
            return

        width = max(self.chart.winfo_width(), 420)
        max_saidas = max(item["Saidas"] for item in resumo) or 1
        row_height = 34
        left = 190
        bar_max = max(width - left - 70, 80)

        for index, item in enumerate(resumo[:40]):
            y = 18 + index * row_height
            label = item["Assessor"][:28]
            bar_width = int((item["Saidas"] / max_saidas) * bar_max)
            self.chart.create_text(10, y + 8, anchor="w", text=label)
            self.chart.create_rectangle(left, y, left + bar_width, y + 18, fill="#2f6fed", width=0)
            self.chart.create_text(
                left + bar_width + 8,
                y + 9,
                anchor="w",
                text=str(item["Saidas"]),
            )

        height = 40 + min(len(resumo), 40) * row_height
        self.chart.configure(scrollregion=(0, 0, width, height))


if __name__ == "__main__":
    app = VazaoApp()
    app.mainloop()
