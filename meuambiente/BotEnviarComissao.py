import tkinter as tk
from tkinter import filedialog
import os
import enviarEmail


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Selecione um Diretório")

        largura_tela = root.winfo_screenwidth() // 2
        altura_tela = root.winfo_screenheight() // 2
        root.geometry(f"{largura_tela}x{altura_tela}")

        self.loading_label = tk.Label(root, text="", wraplength=400, font=("Arial", 12))
        self.loading_label.pack(padx=20, pady=5)

        self.directory_label = tk.Label(root, text="", wraplength=400, font=("Arial", 12))
        self.directory_label.pack(padx=20, pady=5)

        self.select_button = tk.Button(root, text="Selecionar Diretório", command=self.selecionar_diretorio)
        self.select_button.pack(padx=20, pady=5)

        self.execute_button = tk.Button(root, text="Executar", command=self.salvar_e_executar)
        self.execute_button.pack(padx=20, pady=5)

    def selecionar_diretorio(self):
        self.url = filedialog.askdirectory()
        self.directory_label.config(text=self.url)
        if self.url:
            self.directory_label.config(text=self.url)
            self.loading_label.config(text="Diretório selecionado com sucesso!")
        else:
            self.loading_label.config(text="Nenhum diretório selecionado.")


    def salvar_e_executar(self):
        if hasattr(self, 'url') and self.url:  # Verifica se um diretório foi selecionado
            self.loading_label.config(text="Enviando e-mail...")
            enviarEmail.processar_diretorio(self.url)
            self.loading_label.config(text="E-mail enviado com sucesso!")
        else:
            self.loading_label.config(text="Por favor, selecione um diretório primeiro.")

root = tk.Tk()
app = App(root)
root.mainloop()
