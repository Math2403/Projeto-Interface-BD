import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askopenfilename
import mysql.connector
import time
import os

# Função para pegar nome BD
def nome_bd():
    texto_bd = add_nome_bd.get("1.0",END)

    return texto_bd

# Função para pegar nome tabela
def nome_tabela():
    texto_tabela = add_nome_tabela.get("1.0",END)
    return texto_tabela

#Função para pegar nome host
def nome_host():
    texto_host = add_nome_host.get("1.0", END)
    return texto_host

#Função para pegar nome usuário
def usuario():
    texto_usuario = add_nome_usuario.get("1.0", END)
    return texto_usuario

#Função para pegar senha
def senha():
    texto_senha = add_nome_senha.get("1.0", END)
    return texto_senha

# Função criar conexão com o BD
def criar_conexao(colunas=False):
    if os.path.isfile('projetotkinterexcel.xlsx'):
        mysqldb = pd.read_excel('projetotkinterexcel.xlsx', index_col=0)
        if colunas:
            colunas = list(mysqldb.columns)
            return colunas
        else:
            return mysqldb
    else:
        lista = [nome_host(), usuario(), senha()]
        lista1 = [texto.replace('\n', '') for texto in lista]

        mydb = mysql.connector.connect(
            host= lista1[0],
            user = lista1[1],
            password = lista1[2]
        )
        mysqldb = pd.read_sql(f"""SELECT * FROM {nome_bd()}.{nome_tabela()};""", mydb)
        if colunas:
            colunas = list(mysqldb.columns)
            return colunas
        else:
            return mysqldb

# Botão inserir linha
def btn_inserir():

    botao_inserir_valores.place(x = 648, y = 175)

    caixa_inserir1_label.place(x = 500, y = 148)
    caixa_inserir1.place(x = 517, y = 175)

# Função para inserir linha
def btn_inserir_valores():
    
    valor = caixa_inserir1.get("1.0",END)
    valor = valor.split(';')

    if os.path.isfile('projetotkinterexcel.xlsx'):
        copy_mysqldb = pd.read_excel('projetotkinterexcel.xlsx', index_col=0)
    else:
        copy_mysqldb = criar_conexao().copy()

    cols = criar_conexao(colunas=True)

    copy_mysqldb = pd.concat([copy_mysqldb, pd.DataFrame([valor], columns=cols)], ignore_index=True)

    copy_mysqldb.set_index(cols[0], inplace=True)
    label_campo_inserido.place(
        x = 517, y = 210
    )
    copy_mysqldb.to_excel('projetotkinterexcel.xlsx')

#Função ler BD
def btn_ler():
    if os.path.isfile('projetotkinterexcel.xlsx'):
        copy_mysqldb = pd.read_excel('projetotkinterexcel.xlsx', index_col=0)
    else:
        copy_mysqldb = criar_conexao().copy()
    copy_mysqldb.to_excel('projetotkinterexcel.xlsx')
    texto_leitura.pack()
    texto_leitura.place(
        x = 225, y = 340
    )

#Função exibir combobox(lista selecionável)
contador = 0

def lista_colunas():
    global contador
    contador += 1
    if contador == 1:
        global lista_caixa
        lista_caixa = ttk.Combobox(values=criar_conexao(colunas=True))
        lista_caixa_texto = tk.Label(text='Selecionar Coluna', bg='#D9D9D9')
        lista_caixa_texto.place(
        x=500, y = 65,
        width = 200,
        height = 10
        )

        lista_caixa.place(
            x=500, y = 80,
            width = 200,
            height = 30
        )
        return lista_caixa
    else:
        return lista_caixa

#Botão atualizar tabela
def btn_atualizar():
    coluna = lista_colunas().get()
    caixa_texto_label.place(
        x = 500, y = 130
    )
    caixa_texto.pack()
    
    caixa_texto.place(x = 517, y = 150)
    botao_receber_texto.place(x = 648, y = 175)

    caixa_texto_valor.place(x = 517, y = 200)
    caixa_texto_valor_label.place(x = 500, y = 177)

    return coluna

#Funcionalidade botão atualizar tabela
def btn_pegar_id():
    id = caixa_texto.get("1.0",END)
    id = int(id)

    valor = caixa_texto_valor.get("1.0",END)

    if os.path.isfile('projetotkinterexcel.xlsx'):
        copy_mysqldb = pd.read_excel('projetotkinterexcel.xlsx', index_col=0)
    else:
        copy_mysqldb = criar_conexao().copy()
    pos = copy_mysqldb.loc[[id - 1], [btn_atualizar()]] = valor
    label_atualizar_id.place(
        x = 517, y = 250
    )
    texto_leitura.place(
        x = 225, y = 340
    )

    copy_mysqldb.to_excel('projetotkinterexcel.xlsx')

#Botão deletar linha
def btn_deletarlinha():
    botao_dlinha.place(x = 648, y = 203)

    caixa_deletar_linha.place(x = 517, y = 205)
    caixa_deletar_linha_label.place(x = 515, y = 183)

#Funcionalidade botão deletar linha
def deletar_linha():
    if os.path.isfile('projetotkinterexcel.xlsx'):
        copy_mysqldb = pd.read_excel('projetotkinterexcel.xlsx')
    else:
        copy_mysqldb = criar_conexao().copy()
    linha_deletar = caixa_deletar_linha.get("1.0",END)
    linha_deletar = int(linha_deletar)
    copy_mysqldb = copy_mysqldb.loc[copy_mysqldb[copy_mysqldb.columns[0]] != linha_deletar ]
    label_deletar_linha.place(x = 515, y = 230)
    texto_leitura.place(
        x = 225, y = 340
    )
    copy_mysqldb.to_excel('projetotkinterexcel.xlsx')

#Botão deletar coluna
def btn_deletarcoluna():
    coluna = lista_colunas().get()
    botao_dcoluna.place(x = 580, y = 120)

    return coluna

#Funcionalidade botão deletar coluna
def deletar_coluna():
    if os.path.isfile('projetotkinterexcel.xlsx'):
        copy_mysqldb = pd.read_excel('projetotkinterexcel.xlsx', index_col=0)
    else:
        copy_mysqldb = criar_conexao().copy()
    copy_mysqldb = copy_mysqldb.drop(columns=[btn_deletarcoluna()])
    label_deletar_coluna.place(x = 515, y = 164)
    texto_leitura.place(
        x = 225, y = 340
    )
    copy_mysqldb.to_excel('projetotkinterexcel.xlsx')

#Função deletar arquivo
def deletar_arquivo():
    if os.path.isfile('projetotkinterexcel.xlsx'):
        os.remove('projetotkinterexcel.xlsx')
        deletar_arquivo_label.place(x=517, y=300)
    else:
        pass

window = Tk()

#Especificações janela tkinter
window.geometry("710x404")
window.configure(bg = "#ffffff")
canvas = Canvas(
    window,
    bg = "#ffffff",
    height = 404,
    width = 710,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge")
canvas.place(x = 0, y = 0)

background_img = PhotoImage(file = f"background.png")
background = canvas.create_image(
    400.0, 300.0,
    image=background_img)

#Botão deletar arquivo
btn_deletar_arquivo = Button(
    text = 'Deletar arquivo .xlsl',
    command=deletar_arquivo
)

btn_deletar_arquivo.place(
    x = 517, y = 365,
    width = 179,
    height = 30)

deletar_arquivo_label = tk.Label(text='Arquivo deletado com sucesso!', bg='#D9D9D9', fg='#008000')

#Radiobuttons
var_classe = tk.StringVar(value='Nada')
botao_inserir = tk.Radiobutton(text='Inserir Linha', variable=var_classe, value='Inserir', command=btn_inserir, bg='#D9D9D9')
botao_ler = tk.Radiobutton(text='Ler', variable=var_classe, value='Ler', command=btn_ler, bg='#D9D9D9')
botao_atualizar = tk.Radiobutton(text='Atualizar', variable=var_classe, value='Atualizar', command=btn_atualizar, bg='#D9D9D9')
botao_deletarlinha = tk.Radiobutton(text='Deletar Linha', variable=var_classe, value='Deletar Linha', command=btn_deletarlinha, bg='#D9D9D9')
botao_deletarcoluna = tk.Radiobutton(text='Deletar Coluna', variable=var_classe, value='Deletar Coluna', command=btn_deletarcoluna, bg='#D9D9D9')

#Códigos botão atualizar
caixa_texto = tk.Text(width=14, height=1)
caixa_texto_label = tk.Label(text='Digite o id da coluna a ser modificada', bg='#D9D9D9', fg='#FF0000')
texto_recebido = tk.Label(text='', bg='#D9D9D9')
botao_receber_texto = tk.Button(text='Enviar', command=btn_pegar_id)
label_atualizar_id = tk.Label(text='Campo atualizado com sucesso', bg='#D9D9D9', fg='#008000')
caixa_texto_valor = tk.Text(width=14, height=1)
caixa_texto_valor_label = tk.Label(text='Digite o valor', bg='#D9D9D9', fg='#FF0000')

#Códigos botão deletar linha
caixa_deletar_linha = tk.Text(width=14, height=1)
caixa_deletar_linha_label = tk.Label(text='Digite a linha', bg='#D9D9D9', fg='#FF0000')
botao_dlinha = tk.Button(text='Enviar', command=deletar_linha)
linha_deletar = tk.Label(text='', bg='#D9D9D9')
label_deletar_linha = tk.Label(text='Linha deletada com sucesso', bg='#D9D9D9', fg='#008000')

#Códigos botão deletar coluna
botao_dcoluna = tk.Button(text='Enviar', command=deletar_coluna)
label_deletar_coluna = tk.Label(text='Coluna deletada com sucesso', bg='#D9D9D9', fg='#008000')

#Exibir radiobuttons na interface
botao_inserir.place(
    x = 20, y = 12
)
botao_ler.place(
    x = 130, y = 12
)

botao_atualizar.place(
    x = 200, y = 12
)
botao_deletarlinha.place(
    x = 290, y = 12
)
botao_deletarcoluna.place(
    x = 410, y = 12
)

caminho_arquivo = tk.StringVar()

#Exibir mensagem arquivo xlsx criado
texto_leitura = tk.Label(text='Um arquivo .xlsx (Excel) foi criado com a tabela', bg='#D9D9D9', fg='#008000')

#Códigos campo nome banco de dados
add_nome_bd = Text(
    width=14, height=1
    )
add_nome_bd_label = tk.Label(text='Adicionar\n Banco de Dados', bg='#D9D9D9', fg='#FF0000')
add_nome_bd_label.place(x = 20, y=40)

add_nome_bd.place(
    x = 15, y = 78,
    width = 110,
    height = 20)

bt_add_bd = Button(text='Enviar', command=nome_bd)
bt_add_bd.place(x = 22, y = 103, width=100, height=20)
texto_bd = tk.Label(text='', bg='#D9D9D9')

#Códigos campo nome tabela
add_nome_tabela_label = tk.Label(text='Adicionar Tabela', bg='#D9D9D9', fg='#FF0000')
add_nome_tabela_label.place(x = 23, y=125)

add_nome_tabela = Text(
    width=14, height=1
    )

add_nome_tabela.place(
    x = 15, y = 149,
    width = 110,
    height = 20)

bt_add_tabela = Button(text='Enviar', command=nome_tabela)
bt_add_tabela.place(x = 22, y = 173, width=100, height=20)
texto_tabela = tk.Label(text='', bg='#D9D9D9')

#Códigos campo nome host
add_nome_host_label = tk.Label(text='Adicionar Host', bg='#D9D9D9', fg='#FF0000')
add_nome_host_label.place(x = 155, y=55)

add_nome_host = Text(
    width=14, height=1
    )

add_nome_host.place(
    x = 141, y = 78,
    width = 110,
    height = 20)

bt_add_host = Button(text='Enviar', command=nome_host)
bt_add_host.place(x = 147, y = 103, width=100, height=20)
texto_host = tk.Label(text='', bg='#D9D9D9')

#Códigos campo nome usuario
add_nome_usuario_label = tk.Label(text='Adicionar Usuario', bg='#D9D9D9', fg='#FF0000')
add_nome_usuario_label.place(x = 145, y=125)

add_nome_usuario = Text(
    width=14, height=1
    )

add_nome_usuario.place(
    x = 141, y = 149,
    width = 110,
    height = 20)

bt_add_usuario = Button(text='Enviar', command=usuario)
bt_add_usuario.place(x = 147, y = 173, width=100, height=20)
texto_usuario = tk.Label(text='', bg='#D9D9D9')

#Códigos campo senha
add_nome_senha_label = tk.Label(text='Adicionar Senha', bg='#D9D9D9', fg='#FF0000')
add_nome_senha_label.place(x = 90, y=193)

add_nome_senha = Text(
    width=14, height=1
    )

add_nome_senha.place(
    x = 80, y = 217,
    width = 110,
    height = 20)

bt_add_senha = Button(text='Enviar', command=senha)
bt_add_senha.place(x = 85, y = 243, width=100, height=20)
texto_senha = tk.Label(text='', bg='#D9D9D9')

#Códigos inserir linha

botao_inserir_valores = tk.Button(text='Enviar', command=btn_inserir_valores)
caixa_inserir1 = tk.Text(width=14, height=1)
caixa_inserir1_label = tk.Label(text='Digite os valores', bg='#D9D9D9', fg='#FF0000')
label_campo_inserido = tk.Label(text='Linha inserida com sucesso', bg='#D9D9D9', fg='#008000')

#Botão iniciar conexão
bt_iniciar_con = Button(text='Iniciar conexão', command=criar_conexao)
bt_iniciar_con.place(x = 43, y = 270, width=179, height=20)

messagebox.showinfo('Importante!','Após realizar uma operação, feche a interface e a abra novamente!')

window.resizable(False, False)
window.mainloop()

