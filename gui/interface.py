import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from core.processor import process_directory

def launch_app():
    root = tk.Tk()
    root.title("Consolidador de Ordens de Serviço")
    root.geometry("600x420")
    root.resizable(False, False)  # Bloqueia o redimensionamento da janela

    termo_nome = tk.StringVar(value="Linear Construtora LTDA")
    termos_var = tk.StringVar(value="Cancelar, Cancelada, Corrigida")
    pasta_var = tk.StringVar()
    xlsx_dir_var = tk.StringVar()
    xlsx_nome_var = tk.StringVar(value="POs.xlsx")

    def escolher_pasta():
        pasta = filedialog.askdirectory()
        if pasta:
            pasta_var.set(pasta)

    def escolher_saida():
        pasta = filedialog.askdirectory()
        if pasta:
            xlsx_dir_var.set(pasta)

    def iniciar():
        try:
            if not termo_nome.get() or not pasta_var.get() or not xlsx_dir_var.get():
                messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos obrigatórios.")
                return

            btn_iniciar["state"] = "disabled"  # Desabilita o botão ao iniciar
            progress_bar["value"] = 0
            console.delete("1.0", tk.END)
            root.update()  # Atualiza a interface

            # Criar thread para o processamento
            thread = threading.Thread(
                target=executar_processamento,
                daemon=True
            )
            thread.start()

        except Exception as e:
            messagebox.showerror("Erro Inesperado", f"Erro inesperado:\n{str(e)}")
            btn_iniciar["state"] = "normal"  # Reabilita o botão em caso de erro

    def executar_processamento():
        try:
            process_directory(
                base_dir=pasta_var.get(),
                termo_nome=termo_nome.get(),
                termos_exclusao=termos_var.get().split(','),
                xlsx_output_dir=xlsx_dir_var.get(),
                xlsx_filename=xlsx_nome_var.get(),
                update_progress=update_progress,
                log_message=log_message
            )
        except Exception as e:
            log_message(f"[<vermelho>ERRO</vermelho>] {str(e)}")
        finally:
            btn_iniciar["state"] = "normal"


    def update_progress(value, maximum):
        progress_bar["value"] = value
        progress_bar["maximum"] = maximum
        root.update_idletasks()


    def log_message(message):
        import datetime, re
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")

        # Identificar tags <cor>...</cor>
        def insert_colored(text):
            pos = 0
            pattern = re.compile(r"<(.*?)>(.*?)</\1>")
            for match in pattern.finditer(text):
                start, end = match.span()
                if start > pos:
                    console.insert(tk.END, text[pos:start])
                tag, content = match.groups()
                console.insert(tk.END, content, tag)
                pos = end
            if pos < len(text):
                console.insert(tk.END, text[pos:])

        console.insert(tk.END, f"[{timestamp}] ")
        insert_colored(message.strip())
        console.insert(tk.END, "\n")
        console.see(tk.END)

        # Limita a 50 linhas
        if int(console.index('end-1c').split('.')[0]) > 50:
            console.delete("1.0", "2.0")

    # Seção [PDFs]
    frame_pdfs = ttk.LabelFrame(root, text="PDFs", padding=(10, 5))
    frame_pdfs.pack(pady=10, padx=10, fill="x")

    ttk.Label(frame_pdfs, text="Nome do arquivo contém:").grid(row=0, column=0, sticky="w", pady=2)
    ttk.Entry(frame_pdfs, textvariable=termo_nome, width=50).grid(row=0, column=1, sticky="ew", padx=5)

    ttk.Label(frame_pdfs, text="Termos de Exclusão (separar por vírgula):").grid(row=1, column=0, sticky="w", pady=2)
    ttk.Entry(frame_pdfs, textvariable=termos_var, width=50).grid(row=1, column=1, sticky="ew", padx=5)

    ttk.Label(frame_pdfs, text="Pasta Raiz para Varredura:").grid(row=2, column=0, sticky="w", pady=2)
    entry_pasta = ttk.Entry(frame_pdfs, textvariable=pasta_var, width=50)
    entry_pasta.grid(row=2, column=1, sticky="ew", padx=5)
    ttk.Button(frame_pdfs, text="Selecionar", command=escolher_pasta).grid(row=2, column=2, padx=5)

    # Seção [Arquivo Excel]
    frame_excel = ttk.LabelFrame(root, text="Arquivo Excel", padding=(10, 5))
    frame_excel.pack(pady=10, padx=10, fill="x")

    ttk.Label(frame_excel, text="Nome do Arquivo:").grid(row=0, column=0, sticky="w", pady=2)
    ttk.Entry(frame_excel, textvariable=xlsx_nome_var, width=50).grid(row=0, column=1, sticky="ew", padx=5)

    ttk.Label(frame_excel, text="Pasta de Saída:").grid(row=1, column=0, sticky="w", pady=2)
    entry_saida = ttk.Entry(frame_excel, textvariable=xlsx_dir_var, width=50)
    entry_saida.grid(row=1, column=1, sticky="ew", padx=5)
    ttk.Button(frame_excel, text="Selecionar", command=escolher_saida).grid(row=1, column=2, padx=5)

    # Botão de Iniciar
    btn_iniciar = ttk.Button(root, text="Iniciar ", command=iniciar)
    btn_iniciar.pack(pady=10)

    # Barra de progresso
    progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
    progress_bar.pack(pady=5)

    # Console
    ttk.Label(root, text="Console de Progresso:").pack()
    console = tk.Text(root, height=5, width=72)
    console.pack()

    # Configurar tags de cor
    console.tag_config("verde", foreground="green")
    console.tag_config("vermelho", foreground="red")
    console.tag_config("amarelo", foreground="orange")
    console.tag_config("azul", foreground="blue")
    console.tag_config("preto", foreground="black")
    console.tag_config("roxo", foreground="purple")
    console.tag_config("cinza", foreground="gray")

    # Ajuste de responsividade
    for frame in (frame_pdfs, frame_excel):
        frame.grid_columnconfigure(1, weight=1)

    root.mainloop()