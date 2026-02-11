import tkinter as tk

from tkinter import ttk,messagebox

from pymongo import MongoClient

from bson.objectid import ObjectId

import pandas as pd

from datetime import datetime



cliente = MongoClient("mongodb://localhost:27017/")

banco = cliente["CadastroDB"]

colecao = banco["usuarios"]

def carregar_dados(filtros = None):
    
    for item in tree.get_children():
        
        tree.delete(item)

    consulta = filtros if filtros else {} 
    
    registros = colecao.find(consulta)

    for doc in registros:

        tree.insert("","end", values=(str(doc["_id"]), doc["nome"], doc["idade"], doc["email"]))

    atualizar_total()

def atualizar_total():

    total = len(tree.get_children())

    lbl_total.config(text=f"Total de itens: {total}")

def adicionar_dados():
    nome = entrada_nome.get()
    idade = entrada_idade.get()
    email = entrada_email.get()

    if not nome or not idade or not email:
        messagebox.showerror("Erro","Todos os campos são obrigatórios!")
        return
    try:
        idade = int(idade)
    except ValueError:
        messagebox.showerror("Erro","Idade deve ser um número!")
        return

    colecao.insert_one({"nome": nome, "idade": idade, "email": email})
    messagebox.showinfo("Sucesso","Dados cadastrados com sucesso")

    carregar_dados()
    limpar_campos()

def limpar_campos():
    entrada_nome.delete(0, tk.END)
    entrada_idade.delete(0, tk.END)
    entrada_email.delete(0, tk.END)
def alterar_dados():
    
    selecionado = tree.selection()

    if not selecionado:

        messagebox.showerror("Erro","Selecione um registro para alterar!")
        return

    item = tree.item(selecionado)

    doc_item = item["values"][0]

    nome = entrada_nome.get()
    idade = entrada_idade.get()
    email = entrada_email.get()

    if not nome or not idade or not email:

        messagebox.showerror("Erro","Todos os campos são obrigatórios!")

        return

    try:
        idade = int(idade)

    except ValueError:

        messagebox.showerror("Erro","Idade deve ser um número!")

        return

    colecao.update_one({"_id": ObjectId(doc_item)}, {"$set":{"nome": nome, "idade": idade, "email": email}})
    messagebox.showinfo("Sucesso","Registro alterado com sucesso")

    carregar_dados()

    limpar_campos()


def excluir_dados():

    selecionado = tree.selection()

    if not selecionado:

        messagebox.showerror("Erro","Selecione um registro para excluir!")
        return


    item = tree.item(selecionado)

    doc_id = item["values"][0] #id do documento MongoDB

    colecao.delete_one({"_id":ObjectId(doc_id)})
    messagebox.showinfo("Sucesso","Registro excluído com sucesso!")

    carregar_dados() #atualiza o registro

    limpar_campos() 
    

def filtrar_dados():
    nome = entrada_nome.get()
    idade = entrada_idade.get()
    email = entrada_email.get()

    filtros = {}

    if nome:

        filtros["nome"] = {"$regex": nome, "$options":"i"}
    if idade:

        try:

            idade = int(idade)
            
            filtros["idade"] = idade

        except ValueError:

            messagebox.showerror("Erro","Idade deve ser um número!")

            return
        
    if email:

        filtros["email"] = {"$regex": email, "$options":"i"}

    carregar_dados(filtros)
    
def exportar_para_excel():

    #Cria uma lista vazia para armarzenar os registros
    registros = []

    #Itera cada item presente na TreeView, retornando todos os itens (linhas)
    for item in tree.get_children():
        
        registros.append(tree.item(item,"values"))


    
    df = pd.DataFrame(registros, columns=["ID","Nome","Idade","Email"])
    
    carimbo = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    arquivo = f"dados_exportados_{carimbo}.xlsx"  # salva na pasta atual

    df.to_excel(arquivo, index = False) #index = false - diz ao pandas apra não 
    #incluir uma coluna de índice no arquivo Excel

    messagebox.showinfo("Sucesso",f"Dados exportados para o arquivo '{arquivo}' com sucesso")
def selecionar_item(evento):

    selecionado = tree.selection()

    if selecionado:

        item = tree.item(selecionado)

        valores = item["values"]
    
        entrada_nome.delete(0,tk.END)
        entrada_nome.insert(0,valores[1])
        
        entrada_idade.delete(0,tk.END)
        entrada_idade.insert(0,valores[2])
    
        entrada_email.delete(0,tk.END)
        entrada_email.insert(0,valores[3])
    
janela = tk.Tk()

janela.title("Gerenciador de Dados com MongoDB")

janela.geometry("1000x800")

janela.resizable(False,False)

#Frame(quadro) do topo 
frame_topo = tk.Frame(janela, bg="#2a9d8f", height = 50)

frame_topo.pack(fill="x")
lbl_titulo = tk.Label(frame_topo,
                     text="Gerenciador de Dados",
                     bg="#2a9d8f",
                     fg="white",
                     font=("Helvetica",16,"bold"))
lbl_titulo.pack(pady=10)

#Frame(quadro) dos dados (nome,idade e email)
frame_dados = tk.Frame(janela,
                       padx=20,
                       pady=10)

frame_dados.pack(fill="x")

#NOME
tk.Label(frame_dados,
         text="Nome:",
         font=("Helvetica",12)).grid(row=0,column=0,sticky="e")

entrada_nome = tk.Entry(frame_dados, font=("Helvetica",12))
entrada_nome.grid(row=0,column=1,padx=10,pady=5)

#IDADE
tk.Label(frame_dados,
         text="Idade:",
         font=("Helvetica",12)).grid(row=1,column=0,sticky="e")

entrada_idade = tk.Entry(frame_dados, font=("Helvetica",12))

entrada_idade.grid(row=1,column=1,padx=10,pady=5)
#EMAIL
tk.Label(frame_dados,
         text="Email:",
         font=("Helvetica",12)).grid(row=2,column=0,sticky="e")

entrada_email = tk.Entry(frame_dados, font=("Helvetica",12))

entrada_email.grid(row=2,column=1,padx=10,pady=5)

frame_botoes = tk.Frame(janela, pady=10)

frame_botoes.pack(fill="x")

btn_adicionar = tk.Button(frame_botoes,
                         text="Adicionar",
                         command=adicionar_dados,
                         bg="#2a9d8f",
                         fg="white",
                         font=("Helvetica",12),width=10)

btn_adicionar.pack(side="left",padx=10)

btn_alterar = tk.Button(frame_botoes,
                         text="Alterar",
                         command=alterar_dados,
                         bg="#f4a261",
                         fg="white",
                         font=("Helvetica",12),width=10)

btn_alterar.pack(side="left",padx=10)

btn_excluir = tk.Button(frame_botoes,
                         text="Excluir",
                         command=excluir_dados,
                         bg="#e76f51",
                         fg="white",
                         font=("Helvetica",12),width=10)

btn_excluir.pack(side="left",padx=10)

btn_filtrar = tk.Button(frame_botoes,
                         text="Filtrar",
                         command=filtrar_dados,
                         bg="#264653",
                         fg="white",
                         font=("Helvetica",12),width=10)

btn_filtrar.pack(side="left",padx=10)
btn_exportar = tk.Button(frame_botoes,
                         text="Exportar para Excel",
                         command=exportar_para_excel,
                         bg="#6a994e",
                         fg="white",
                         font=("Helvetica",12),width=15)

btn_exportar.pack(side="left",padx=10)

lbl_total = tk.Label(janela,
                    text="Total de itens: 0",
                    font=("Helvetica",12),
                    bg="#f0f0f0")
lbl_total.pack()

frame_lista = tk.Frame(janela,pady=20)

frame_lista.pack(fill="both", expand=True)

estilo = ttk.Style()

estilo.configure("Treeview", font=("Helvetica",12))

estilo.configure("Treeview.Heading", font=("Helvetica",12,"bold"))

tree = ttk.Treeview(frame_lista,
                   columns=("ID", "Nome","Idade","Email"),
                    show = "headings",
                    style = "Treeview")

scroll_y = ttk.Scrollbar(frame_lista,orient="vertical",command = tree.yview)

tree.configure(yscrollcommand=scroll_y.set)

scroll_y.pack(side="right", fill="y")

tree.pack(fill="both", expand=True)

tree.heading("ID", text="ID")

tree.heading("Nome", text="Nome")

tree.heading("Idade", text="Idade")
    
tree.heading("Email", text="Email")

tree.column("ID", width=170,anchor="center")
tree.column("Nome", width=200,anchor="center")
tree.column("Idade", width=100,anchor="center")
tree.column("Email", width=200,anchor="w")

tree.bind("<ButtonRelease-1>", selecionar_item)

carregar_dados()


janela.mainloop()