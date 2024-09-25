import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import os

# Função para criar conexão com o banco de dados SQLite
def criar_conexao(db_file):
    conn = None
    try:
        conn = sqlite3.connect(db_file)
        return conn
    except sqlite3.Error as e:
        messagebox.showerror("Erro de Conexão", str(e))
    return conn

# Função para criar a tabela de produtos
def criar_tabela(conn):
    try:
        sql_create_produtos_table = """
        CREATE TABLE IF NOT EXISTS produtos (
            id integer PRIMARY KEY,
            produto text NOT NULL,
            setor text NOT NULL,
            lancamento text NOT NULL,
            qualificacao integer NOT NULL,
            treinamento integer NOT NULL,
            manual_datasheet integer NOT NULL,
            laboratorio integer NOT NULL
        );
        """
        conn.execute(sql_create_produtos_table)
    except sqlite3.Error as e:
        messagebox.showerror("Erro de Criação", str(e))

# Função para consultar dados
def consultar_dados(conn):
    sql_select = "SELECT * FROM produtos"
    cur = conn.cursor()
    cur.execute(sql_select)
    rows = cur.fetchall()
    df = pd.DataFrame(rows, columns=['id', 'produto', 'setor', 'lancamento', 'qualificacao', 'treinamento', 'manual_datasheet', 'laboratorio'])
    return df

# Função para calcular a qualificação com base em treinamento, manual/datasheet e laboratório
def calcular_qualificacao(treinamento, manual_datasheet, laboratorio):
    # Pesos: treinamento (3), manual/datasheet (1), laboratório (2)
    peso_total = 3 + 1 + 2
    qualificacao = (treinamento * 3 + manual_datasheet * 1 + laboratorio * 2) / peso_total
    return round(qualificacao)

# Função para inserir dados
def inserir_dados(conn, produto, setor, lancamento, treinamento, manual_datasheet, laboratorio):
    qualificacao = calcular_qualificacao(treinamento, manual_datasheet, laboratorio)
    
    sql_insert = """ INSERT INTO produtos(produto, setor, lancamento, qualificacao, treinamento, manual_datasheet, laboratorio)
                     VALUES(?, ?, ?, ?, ?, ?, ?) """
    cur = conn.cursor()
    cur.execute(sql_insert, (produto, setor, lancamento, qualificacao, treinamento, manual_datasheet, laboratorio))
    conn.commit()

# Função para editar dados
def editar_dados(conn, id, produto, setor, lancamento, treinamento, manual_datasheet, laboratorio):
    qualificacao = calcular_qualificacao(treinamento, manual_datasheet, laboratorio)
    
    sql_update = """ UPDATE produtos
                     SET produto = ?, setor = ?, lancamento = ?, qualificacao = ?, treinamento = ?, manual_datasheet = ?, laboratorio = ?
                     WHERE id = ? """
    cur = conn.cursor()
    cur.execute(sql_update, (produto, setor, lancamento, qualificacao, treinamento, manual_datasheet, laboratorio, id))
    conn.commit()

# Função para exportar dados para XLSX
def exportar_para_xlsx(df):
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Sucesso", f"Dados exportados para {file_path}")

# Função para capturar dados do usuário
def capturar_dados_do_usuario(conn, produto_entry, setor_entry, lancamento_entry, treinamento_entry, manual_entry, laboratorio_entry):
    produto = produto_entry.get()
    setor = setor_entry.get()
    lancamento = lancamento_entry.get()

    try:
        treinamento = int(treinamento_entry.get())
        manual_datasheet = int(manual_entry.get())
        laboratorio = int(laboratorio_entry.get())
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira valores numéricos para treinamento, manual/datasheet e laboratório.")
        return

    try:
        pd.to_datetime(lancamento, format='%Y-%m-%d')
    except ValueError:
        messagebox.showerror("Erro", "Data inválida. Use o formato AAAA-MM-DD.")
        return

    inserir_dados(conn, produto, setor, lancamento, treinamento, manual_datasheet, laboratorio)
    messagebox.showinfo("Sucesso", "Dados inseridos com sucesso!")

# Função para monitorar o banco de dados e atualizar automaticamente os dados
def monitorar_e_atualizar(conn, tree):
    rows = consultar_dados(conn).values.tolist()
    
    for i in tree.get_children():
        tree.delete(i)

    for row in rows:
        tree.insert("", tk.END, values=row)

    tree.after(5000, lambda: monitorar_e_atualizar(conn, tree))  # Atualiza a cada 5 segundos

# Função para abrir a janela de edição
def abrir_janela_edicao(conn, tree):
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        values = item['values']
        id_selecionado = values[0]

        # Criar nova janela para edição
        janela_edicao = tk.Toplevel()
        janela_edicao.title("Editar Produto")

        tk.Label(janela_edicao, text="Produto").grid(row=0, column=0, padx=5, pady=5)
        produto_entry = tk.Entry(janela_edicao)
        produto_entry.grid(row=0, column=1, padx=5, pady=5)
        produto_entry.insert(0, values[1])

        tk.Label(janela_edicao, text="Setor").grid(row=1, column=0, padx=5, pady=5)
        setor_entry = tk.Entry(janela_edicao)
        setor_entry.grid(row=1, column=1, padx=5, pady=5)
        setor_entry.insert(0, values[2])

        tk.Label(janela_edicao, text="Lançamento (AAAA-MM-DD)").grid(row=2, column=0, padx=5, pady=5)
        lancamento_entry = tk.Entry(janela_edicao)
        lancamento_entry.grid(row=2, column=1, padx=5, pady=5)
        lancamento_entry.insert(0, values[3])

        tk.Label(janela_edicao, text="Treinamento").grid(row=3, column=0, padx=5, pady=5)
        treinamento_entry = tk.Entry(janela_edicao)
        treinamento_entry.grid(row=3, column=1, padx=5, pady=5)
        treinamento_entry.insert(0, values[5])

        tk.Label(janela_edicao, text="Manual/Datasheet").grid(row=4, column=0, padx=5, pady=5)
        manual_entry = tk.Entry(janela_edicao)
        manual_entry.grid(row=4, column=1, padx=5, pady=5)
        manual_entry.insert(0, values[6])

        tk.Label(janela_edicao, text="Laboratório").grid(row=5, column=0, padx=5, pady=5)
        laboratorio_entry = tk.Entry(janela_edicao)
        laboratorio_entry.grid(row=5, column=1, padx=5, pady=5)
        laboratorio_entry.insert(0, values[7])

        # Botão para salvar a edição
        salvar_btn = tk.Button(janela_edicao, text="Salvar", command=lambda: salvar_edicao(conn, tree, janela_edicao, id_selecionado, produto_entry, setor_entry, lancamento_entry, treinamento_entry, manual_entry, laboratorio_entry))
        salvar_btn.grid(row=6, column=0, columnspan=2, pady=10)

# Função para salvar a edição
def salvar_edicao(conn, tree, janela_edicao, id, produto_entry, setor_entry, lancamento_entry, treinamento_entry, manual_entry, laboratorio_entry):
    produto = produto_entry.get()
    setor = setor_entry.get()
    lancamento = lancamento_entry.get()

    try:
        treinamento = int(treinamento_entry.get())
        manual_datasheet = int(manual_entry.get())
        laboratorio = int(laboratorio_entry.get())
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira valores numéricos para treinamento, manual/datasheet e laboratório.")
        return

    editar_dados(conn, id, produto, setor, lancamento, treinamento, manual_datasheet, laboratorio)
    monitorar_e_atualizar(conn, tree)
    janela_edicao.destroy()  # Fecha a janela de edição após salvar

# Função para excluir o item selecionado
def excluir_item(conn, tree):
    selected_item = tree.selection()
    if selected_item:
        item = tree.item(selected_item)
        id = item['values'][0]
        
        try:
            cur = conn.cursor()
            cur.execute("DELETE FROM produtos WHERE id = ?", (id,))
            conn.commit()
            monitorar_e_atualizar(conn, tree)
            messagebox.showinfo("Sucesso", "Linha excluída com sucesso!")
        except sqlite3.Error as e:
            messagebox.showerror("Erro ao excluir linha", str(e))

# Função principal da interface
def main():
    # Criar conexão com o banco de dados
    database = "produtos_intelbras.db"
    conn = criar_conexao(database)

    if conn is not None:
        criar_tabela(conn)

        root = tk.Tk()
        root.title("Gerenciador de Produtos")

        frame = tk.Frame(root)
        frame.pack(padx=10, pady=10)

        tk.Label(frame, text="Produto").grid(row=0, column=0, padx=5, pady=5)
        produto_entry = tk.Entry(frame)
        produto_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Label(frame, text="Setor").grid(row=1, column=0, padx=5, pady=5)
        setor_entry = tk.Entry(frame)
        setor_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(frame, text="Lançamento (AAAA-MM-DD)").grid(row=2, column=0, padx=5, pady=5)
        lancamento_entry = tk.Entry(frame)
        lancamento_entry.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(frame, text="Treinamento").grid(row=3, column=0, padx=5, pady=5)
        treinamento_entry = tk.Entry(frame)
        treinamento_entry.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(frame, text="Manual/Datasheet").grid(row=4, column=0, padx=5, pady=5)
        manual_entry = tk.Entry(frame)
        manual_entry.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(frame, text="Laboratório").grid(row=5, column=0, padx=5, pady=5)
        laboratorio_entry = tk.Entry(frame)
        laboratorio_entry.grid(row=5, column=1, padx=5, pady=5)

        inserir_btn = tk.Button(frame, text="Inserir Dados", command=lambda: capturar_dados_do_usuario(conn, produto_entry, setor_entry, lancamento_entry, treinamento_entry, manual_entry, laboratorio_entry))
        inserir_btn.grid(row=6, column=0, columnspan=2, pady=10)

        exportar_btn = tk.Button(frame, text="Exportar para XLSX", command=lambda: exportar_para_xlsx(consultar_dados(conn)))
        exportar_btn.grid(row=7, column=0, columnspan=2, pady=10)

        # Configurar Treeview para monitoramento
        colunas = ("ID", "Produto", "Setor", "Lançamento", "Qualificação", "Treinamento", "Manual/Datasheet", "Laboratório")
        tree = ttk.Treeview(root, columns=colunas, show="headings")
        for col in colunas:
            tree.heading(col, text=col)
        tree.pack(fill=tk.BOTH, expand=True)

        monitorar_e_atualizar(conn, tree)

        # Botões para editar e excluir
        editar_btn = tk.Button(frame, text="Editar", command=lambda: abrir_janela_edicao(conn, tree))
        editar_btn.grid(row=8, column=0, pady=10)

        excluir_btn = tk.Button(frame, text="Excluir", command=lambda: excluir_item(conn, tree))
        excluir_btn.grid(row=8, column=1, pady=10)

        root.mainloop()

if __name__ == '__main__':
    main()
