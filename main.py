from gui.interface import launch_app
from core.ocr_utils import check_dependencies
import tkinter as tk
from tkinter import messagebox

if __name__ == "__main__":
    errors, poppler_path = check_dependencies()

    if errors:
        root = tk.Tk()
        root.withdraw()

        # Formatar a mensagem de erro
        error_message = "Algumas dependências não foram encontradas:\n\n" + "\n".join(errors)
        error_message += "\n\nAlguns recursos podem não funcionar corretamente."

        messagebox.showwarning(
            "Dependências Ausentes",
            error_message
        )
        root.destroy()

    launch_app()