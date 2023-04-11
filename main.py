import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import Calendar, DateEntry
import os
import openpyxl

# Função para permitir somente números na Entry
def numero(input):
    if input.isdigit() or input == "":
        print(input)
        return True
    else:
        print(input)
        return False

# Função botão de dados
def salvar_data():
        nome = nome_entry.get()
        sobrenome = sobrenome_entry.get()
        cpf = cpf_entry.get()
        datanasc = dataNasc_entry.get()
        idunico = idUnico_entry.get()
        atendimento = atendimento_entry.get()
        datapront = dataPront_entry.get()
        estante = estante_combobox.get()
        prateleira = prateleira_combobox.get()
        setor = setor_combobox.get()
        local = local_combobox.get()
        caixa = caixa_entry.get()
        status = instru_label
        if nome and sobrenome and cpf and idunico and atendimento and local:
            filepath = "data.xlsx"
            if not os.path.exists(filepath): 
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["Registro", "Atendimento", "Nome", "Sobrenome", "CPF", "Data do Prontuário", "Local", "Estante", "Prateleira", "Caixa", "Setor", "Data de Nascimento"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([idunico, atendimento, nome, sobrenome, cpf, datapront, local, estante, prateleira, caixa, setor, datanasc])
            workbook.save(filepath)


            status = instru_label.config(text="Registro de " + nome + " salvo.")
            nome = nome_entry.delete(0, tk.END)
            sobrenome = sobrenome_entry.delete(0, tk.END)
            cpf = cpf_entry.delete(0, tk.END)
            datanasc = dataNasc_entry.delete(0, tk.END)
            idunico = idUnico_entry.delete(0, tk.END)
            atendimento = atendimento_entry.delete(0, tk.END)
            datapront = dataPront_entry.delete(0, tk.END)
            estante = estante_combobox.delete(0, tk.END)
            prateleira = prateleira_combobox.delete(0, tk.END)
            setor = setor_combobox.delete(0, tk.END)
            local = local_combobox.delete(0, tk.END)
            caixa = caixa_entry.delete(0, tk.END)


        else:
            tk.messagebox.showwarning(title="Erro", message="Os campos Nome, Sobrenome, CPF, Registro, Atendimento e Local devem ser preenchidos.")


# Root
root = tk.Tk()
root.iconbitmap("data/bar.ico")
root.title("GUISAME")
root.option_add("*tearOff", False)
style = ttk.Style(root)
root.tk.call("source", "data/tema/forest-dark.tcl")
ttk.Style().theme_use("forest-dark")
enumero=root.register(numero)

# Responsividade
root.columnconfigure(index=0, weight=1)
root.columnconfigure(index=1, weight=1)
root.columnconfigure(index=2, weight=1)
root.rowconfigure(index=0, weight=1)
root.rowconfigure(index=1, weight=1)
root.rowconfigure(index=2, weight=1)

# Frame Principal
info_pac_frame = ttk.Frame(root)
info_pac_frame.pack(fill="both", expand=True)
info_pac_frame.grid(row=0, column=0, padx=20, pady=10, sticky="news")

nome_label = ttk.Label(info_pac_frame, text="Nome",  width=15)
nome_label.grid(row=0, column=0)
nome_entry = ttk.Entry(info_pac_frame,  width=25)
nome_entry.grid(row=0, column=1)

sobrenome_label = ttk.Label(info_pac_frame, text="Sobrenome",  width=15)
sobrenome_label.grid(row=0, column=2)
sobrenome_entry = ttk.Entry(info_pac_frame,  width=25)
sobrenome_entry.grid(row=0, column=3)

cpf_label = ttk.Label(info_pac_frame, text="CPF")
cpf_label.grid(row=1, column=0)
cpf_entry = ttk.Entry(info_pac_frame, validate="key", validatecommand=(enumero, '%P'))
cpf_entry.grid(row=1, column=1)

dataNasc_label = ttk.Label(info_pac_frame, text="Data de Nascimento")
dataNasc_label.grid(row=1, column=2)
dataNasc_entry = DateEntry(info_pac_frame, background="green", date_pattern="dd/mm/yyyy")
dataNasc_entry.grid(row=1, column=3)

idUnico_label = ttk.Label(info_pac_frame, text="Registro Único")
idUnico_label.grid(row=2, column=0)
idUnico_entry = ttk.Entry(info_pac_frame, validate="key", validatecommand=(enumero, '%P'))
idUnico_entry.grid(row=2, column=1)

atendimento_label = ttk.Label(info_pac_frame, text="Atendimento")
atendimento_label.grid(row=2, column=2)
atendimento_entry = ttk.Entry(info_pac_frame, validate="key", validatecommand=(enumero, '%P'))
atendimento_entry.grid(row=2, column=3)

dataPront_label = ttk.Label(info_pac_frame, text="Data do Prontuário")
dataPront_label.grid(row=3, column=0)
dataPront_entry = DateEntry(info_pac_frame, background="green", date_pattern="dd/mm/yyyy")
dataPront_entry.grid(row=3, column=1)

setor_label = ttk.Label(info_pac_frame, text="Setor")
setor_label.grid(row=3, column=2)
setor_combobox = ttk.Combobox(info_pac_frame, values=["DPT", "DDQ", "HD"])
setor_combobox.grid(row=3, column=3)

estante_label = ttk.Label(info_pac_frame, text="Estante")
estante_label.grid(row=4, column=0)
estante_combobox = ttk.Combobox(info_pac_frame, values=["A","B","C","D","F","G","H","I"])
estante_combobox.grid(row=4, column=1)

prateleira_label = ttk.Label(info_pac_frame, text="Prateleira")
prateleira_label.grid(row=4, column=2)
prateleira_combobox = ttk.Combobox(info_pac_frame, values=["1","2","3","4","5","6","7","8","9"])
prateleira_combobox.grid(row=4, column=3)

local_label = ttk.Label(info_pac_frame, text="Local")
local_label.grid(row=5, column=0)
local_combobox = ttk.Combobox(info_pac_frame, values=["Hospital","Externo","Morto"])
local_combobox.grid(row=5, column=1)

caixa_label = ttk.Label(info_pac_frame, text="Nº Caixa")
caixa_label.grid(row=5, column=2)
caixa_entry = ttk.Entry(info_pac_frame)
caixa_entry.grid(row=5, column=3)

instru_label = ttk.Label(info_pac_frame, text="")
instru_label.grid(row=6, column=0, columnspan=2)


# Botão final
button = ttk.Button(info_pac_frame, text="Salvar", style='Accent.TButton', width=25, command=salvar_data)
button.grid(row=6, column=2, columnspan=2, sticky="news", padx=20, pady=20)
root.bind('<Return>', lambda event=None: button.invoke())

# Função para configurar o espaço de todos itens
for widget in info_pac_frame.winfo_children():
    widget.grid_configure(padx=10, pady=10, sticky="news")


root.mainloop()