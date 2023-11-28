import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog,messagebox
import os
from tkinter import ttk


import warnings
warnings.filterwarnings('ignore')


def obter_esquema_catalogo(arquivo_xml, catalogo_escolhido):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    esquema_catalogo = {}

    for catalogo in root.findall(f'.//Catalog[@Name="{catalogo_escolhido}"]'):
        for campo in catalogo.findall('.//Field'):
            nome_campo = campo.get('Name')
            tamanho = int(campo.get('Length'))
            esquema_catalogo[nome_campo] = tamanho

    return esquema_catalogo

def obter_nomes_catalogos(arquivo_xml):
    tree = ET.parse(arquivo_xml)
    root = tree.getroot()

    nomes_catalogos = []

    for catalogo in root.findall('.//Catalog'):
        nome_catalogo = catalogo.get('Name')
        if nome_catalogo:
            nomes_catalogos.append(nome_catalogo)

    return nomes_catalogos


def ler_ff(baseff,esquema_catalogo):
    base = pd.read_csv(baseff,header=None)
    base[0]=base[0].fillna(sum(esquema_catalogo.values()))

    for nome,carac in esquema_catalogo.items():
        base[nome]=base[0].apply(lambda x:x[:carac])
        base[0] = base[0].apply(lambda x:x[carac:])
    
    base.drop(0,axis=1,inplace=True)
    return base


def seleciona_caminho():
    global caminho_selecionado
    caminho_selecionado = filedialog.askdirectory(title="Selecione a pasta onde será salvo o arquivo tratado")
    caminho.config(text=f"{caminho_selecionado}")
    


def seleciona_ff():
    global baseff_selecionado
    baseff_selecionado = filedialog.askopenfilename(title="Selecione o arquivo base FF", filetypes=[(".txt", "*.txt")])
    baseff.config(text=f"{os.path.basename(baseff_selecionado)}")
    

def seleciona_xml():
    global xml_selecionado
    xml_selecionado = filedialog.askopenfilename(title="Selecione o arquivo XML", filetypes=[("XML files", "*.xml")])
    xml.config(text=f"{os.path.basename(xml_selecionado)}")

    nomes_catalogos = obter_nomes_catalogos(xml_selecionado)
    catalogo['values']=nomes_catalogos
    catalogo.set("Selecione o catálogo")

def novo_flat_file():
    global mensagem_sucesso
    caminho.config(text="Local")
    baseff.config(text="Arquivo Flat File")
    xml.config(text="Arquivo de Catálogos")
    catalogo.delete('0','end')
    catalogo.insert(0,"CatalogName")
    info_text.delete(1.0,tk.END)
    button_salvar.config(state='disabled')

    try:
        mensagem_sucesso.destroy()
    except:
        pass

def finalizar():
    root.destroy()


def exibe_mensagem_sucesso():
    global mensagem_sucesso
    mensagem_sucesso = tk.Toplevel(root)
    mensagem_sucesso.title("Sucesso!")

    label = ttk.Label(mensagem_sucesso,text=f"Arquivo salvo em:\n{caminho_selecionado}")
    label.pack(padx=10,pady=10)

    button_novo_ff = ttk.Button(mensagem_sucesso,text="Novo Flat File",command=novo_flat_file)
    button_novo_ff.pack(side='left',padx=5,pady=5)

    button_encerrar = ttk.Button(mensagem_sucesso,text="Encerrar",command=finalizar)
    button_encerrar.pack(side='left',padx=5,pady=5)


def salvar():
    if caminho_selecionado and baseff_selecionado and xml_selecionado:
        esquema_catalogo = obter_esquema_catalogo(xml_selecionado,catalogo.get())
        df_tratado = ler_ff(baseff_selecionado,esquema_catalogo)

        nome_padrao =f"{os.path.basename(baseff_selecionado)}_tratada"

        caminho_arquivo = filedialog.asksaveasfile(title="Salvar Arquivo Tratado",
                                                   filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                                                   defaultextension=".xlsx",
                                                   initialfile=nome_padrao,
                                                   initialdir=caminho_selecionado)
        

        if not caminho_arquivo:
            return
        
        formato_selecionado = caminho_arquivo.name.split('.')[-1].lower()
        if formato_selecionado=="xlsx":
            df_tratado.to_excel(caminho_arquivo.name,index=False)
        
        elif formato_selecionado=="csv":
            df_tratado.to_csv(caminho_arquivo.name,sep=";",decimal=",",index=False)
        else:
            messagebox.showerror("Erro!","Formato de arquivo inválido")
            return
        
        exibe_mensagem_sucesso()

def executar():
    if caminho_selecionado and baseff_selecionado and xml_selecionado:
        info_text.delete(1.0,tk.END)
        info_text.insert(tk.END, f"Caminho: {caminho_selecionado}\n")
        info_text.insert(tk.END, f"Arquivo Flat File: {baseff_selecionado}\n")
        info_text.insert(tk.END, f"Arquivo de Catálogos: {xml_selecionado}\n")
        info_text.insert(tk.END, f"Catalogo: {catalogo.get()}\n\n")

        #Ativar Botao "Salvar Arquivo Tratado"
        button_salvar.config(state="normal")



root = tk.Tk()
style = ttk.Style(root)
root.title("Tratamento de Flat File")
style.theme_use('xpnative')

frame = ttk.Frame(root)
frame.pack()

#Input Frame
widgets_frame = ttk.LabelFrame(frame,text="Dados do Flat File",width=300,height=200)
widgets_frame.grid(row=0,column=0,padx=20,pady=20)


caminho = ttk.Button(widgets_frame,text="Local",command=seleciona_caminho,width=55)
caminho.grid(row=0,column=0,padx=5,pady=5,sticky="nsew",columnspan=2)

baseff = ttk.Button(widgets_frame,text="Arquivo Flat File",command=seleciona_ff)
baseff.grid(row=1,column=0,padx=5,pady=5,sticky="nsew",columnspan=2)

xml = ttk.Button(widgets_frame,text="Arquivo de Catálogos",command=seleciona_xml)
xml.grid(row=2,column=0,padx=5,pady=5,sticky="nsew",columnspan=2)

catalogo = ttk.Combobox(widgets_frame)
catalogo.grid(row=3,column=0,padx=5,pady=5,sticky="nsew",columnspan=2)


button_execute = ttk.Button(widgets_frame, text="Tratar Flat File", command=executar)
button_execute.grid(row=4, column=0, padx=5, pady=5, sticky="nsew")

button_novo_ff = ttk.Button(widgets_frame,text="Apagar Dados",command=novo_flat_file)
button_novo_ff.grid(row=4,column=1,padx=5,pady=5,sticky='nsew')

#Status Frame
statusFrame = ttk.Frame(frame)
statusFrame.grid(row=0,column=1,pady=20,padx=50)
statusScroll = ttk.Scrollbar(statusFrame)
statusScroll.pack(side="right",fill="y")

info_text = tk.Text(statusFrame, wrap="word", yscrollcommand=statusScroll.set, height=10, width=50)
info_text.pack(expand=True, fill="both")
statusScroll.config(command=info_text.yview)

button_salvar = ttk.Button(statusFrame,text="Salvar Arquivo Tratado",command=salvar,state='disabled')
button_salvar.pack(side="bottom", pady=5)



root.mainloop()