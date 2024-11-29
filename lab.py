import serial
import datetime
import time
import csv

import tkinter as tk
from tkinter.ttk import *
import sqlite3
from tkinter import messagebox
from tkinter import *
from tkinter import ttk
from reportlab.pdfgen import canvas

class VidrosApp:

    first = True
    running = False
    ensaio_id = 0
    f_valor_lido = 0.0
    f_valor_maximo = 0.0
    f_valor_minimo = 0.0

    def __init__(self, root):
        self.root = root
        self.root.title("Falcão Bauer")

        # # Create a database or connect to an existing one
        self.conn = sqlite3.connect("labvidros.db")
        self.cursor = self.conn.cursor()


        # # Create a table if it doesn't exist
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS empresas (
            empresa TEXT
        )''')
        self.conn.commit()
        
        # # Create a table if it doesn't exist
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS modelos (
            modelo TEXT
        )''')
        self.conn.commit()

        # # Create a table if it doesn't exist
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS ensaios_summary (
            id INTEGER PRIMARY KEY,
            timestampInicio TIMESTAMP,
            empresa TEXT,
            modelo TEXT,
            tamanho TEXT,
            numero TEXT,
            processo TEXT,
            valorMinimo REAL,
            valorMaximo REAL,
            timestampFim TIMESTAMP
        )''')
        self.conn.commit()

        # # Create a table if it doesn't exist
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS ensaios_values (
            id_summary INTEGER,
            value REAL
        )''')
        self.conn.commit()

        # Create GUI elements
        self.ensaio_label = tk.Label(root, text="Abrir Ensaio:")
        self.ensaio_label.pack()

        self.comboEnsaios = ttk.Combobox(root, values=cb_ensaios)
        self.comboEnsaios.pack()

        self.open_ensaios_button = tk.Button(root, text="Abrir", command=self.load_ensaio)
        self.open_ensaios_button.pack()

        self.empresa_label = tk.Label(root, text="Empresa:")
        self.empresa_label.pack()

        self.comboEmpresas = ttk.Combobox(root, values=cb_empresas)
        self.comboEmpresas.pack()

        self.cad_empresas_button = tk.Button(root, text="Gerenciar Empresas", command=self.wnd_cadastro_empresas)
        self.cad_empresas_button.pack()

        self.modelo_label = tk.Label(root, text="Modelo da Amostra:")
        self.modelo_label.pack()

        self.comboModelos = ttk.Combobox(root, values=cb_modelos)
        self.comboModelos.pack()

        self.cad_modelos_button = tk.Button(root, text="Gerenciar Modelos", command=self.wnd_cadastro_modelos)
        self.cad_modelos_button.pack()

        self.load_combo_ensaios()
        self.load_combo_empresas()
        self.load_combo_modelos()

        self.tamanho_label = tk.Label(root, text="Tamanho da Amostra:")
        self.tamanho_label.pack()

        self.tamanho_entry = tk.Entry(root)
        self.tamanho_entry.pack()

        self.numero_label = tk.Label(root, text="Número da Amostra:")
        self.numero_label.pack()

        self.numero_entry = tk.Entry(root)
        self.numero_entry.pack()

        self.processo_label = tk.Label(root, text="Processo:")
        self.processo_label.pack()

        self.processo_entry = tk.Entry(root)
        self.processo_entry.pack()

        self.init_button = tk.Button(root, text="Iniciar", command=self.iniciar)
        self.init_button.pack()

        self.lido_label = tk.Label(root, text="Valor lido:")
        self.lido_label.pack()

        self.lido_entry = tk.Entry(root, textvariable=read_value, font=("DSeg",48), fg=("green"), bg=("black"))
        self.lido_entry.pack()

        self.maior_label = tk.Label(root, text="Maior Valor:")
        self.maior_label.pack()

        self.maior_entry = tk.Entry(root, textvariable=max_value, font=("DSeg",48), fg=("blue"), bg=("black"))
        self.maior_entry.pack()

        self.menor_label = tk.Label(root, text="Menor Valor:")
        self.menor_label.pack()

        self.menor_entry = tk.Entry(root, textvariable=min_value, font=("DSeg",48), fg=("green"), bg=("black"))
        self.menor_entry.pack()

        self.parar_button = tk.Button(root, text="Parar", command=self.parar_ensaio)
        self.parar_button.pack()

        self.imprimir_button = tk.Button(root, text="Imprimir", command=self.GeneratePDF)
        self.imprimir_button.pack()

    def GeneratePDF(self):
        try:
            global timestampInicio
            tsIni = timestampInicio.replace(':','-')
            nome_pdf = 'Ensaio ' + tsIni
            pdf = canvas.Canvas('{}.pdf'.format(nome_pdf))
            #x = 720
            #for nome,idade in lista.items():
            #    x -= 20
            #    pdf.drawString(247,x, '{} : {}'.format(nome,idade))
            pdf.setTitle(nome_pdf)
            pdf.setFont("Helvetica-Oblique", 14)
            pdf.drawString(145,790, nome_pdf)
            pdf.setFont("Helvetica-Oblique", 14)
            pdf.drawString(145,740, 'Resultados')
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(145,680, 'Empresa: ' + self.comboEmpresas.get())
            pdf.drawString(145,640, 'Modelo: ' + self.comboModelos.get())
            pdf.drawString(145,600, 'Tamanho: ' + self.tamanho_entry.get())
            pdf.drawString(145,560, 'Número: ' + self.numero_entry.get())
            pdf.drawString(145,520, 'Processo: ' + self.processo_entry.get())
            pdf.drawString(145,480, 'Valor Mínimo: ' + self.menor_entry.get())
            pdf.drawString(145,440, 'Valor Máximo: ' + self.maior_entry.get())
            pdf.save()
            messagebox.showwarning("AVISO", '{}.pdf criado com sucesso!'.format(nome_pdf))
        except:
            messagebox.showwarning("ERRO",'Erro ao gerar {}.pdf'.format(nome_pdf))

    # função que permite abrir a janela de cadastro de Empresas
    def wnd_cadastro_empresas(self):
        # janelas adicionais são criadas a partir da classe Toplevel
        wnd_cad_empresas = Toplevel()
        
        # define as dimensões da janela de cadastro de empresas
        wnd_cad_empresas.geometry("350x350")
        
        # define o título da janela de cadastro de empresas
        wnd_cad_empresas.title("Cadastro de Empresas")
        
        wnd_cad_empresas.empresa_label = tk.Label(wnd_cad_empresas, text="Empresa:")
        wnd_cad_empresas.empresa_label.pack()

        wnd_cad_empresas.empresa_entry = tk.Entry(wnd_cad_empresas, width=30)
        wnd_cad_empresas.empresa_entry.pack()

        wnd_cad_empresas.scrollbar = tk.Scrollbar(wnd_cad_empresas)
        #scrollbar.pack(side=RIGHT, fill=Y)
        wnd_cad_empresas.scrollbar.pack(side="right", fill="y")

        wnd_cad_empresas.empresas_listbox = tk.Listbox(wnd_cad_empresas, width=40,yscrollcommand=wnd_cad_empresas.scrollbar.set)
        wnd_cad_empresas.empresas_listbox.pack()

        wnd_cad_empresas.add_button = tk.Button(wnd_cad_empresas, text="Adicionar", command=lambda: self.add_empresa(wnd_cad_empresas))
        wnd_cad_empresas.add_button.pack()

        wnd_cad_empresas.scrollbar.config(command=wnd_cad_empresas.empresas_listbox.yview)

        self.load_empresas(wnd_cad_empresas)
        
        wnd_cad_empresas.update_button = tk.Button(wnd_cad_empresas, text="Update", command=lambda: self.update_empresa(wnd_cad_empresas))
        wnd_cad_empresas.update_button.pack()

        wnd_cad_empresas.delete_button = tk.Button(wnd_cad_empresas, text="Delete", command=lambda: self.delete_empresa(wnd_cad_empresas))
        wnd_cad_empresas.delete_button.pack()

        # coloca o foco na janela de cadastro de empresas
        wnd_cad_empresas.focus_set()
        # desabilita a janela principal enquanto a janela cadastro de empresas estiver aberta
        wnd_cad_empresas.grab_set()
        # inicia o loop de eventos da janela cadastro de empresas
        wnd_cad_empresas.mainloop()

    # função que permite abrir a janela de cadastro de Modelos
    def wnd_cadastro_modelos(self):
        # janelas adicionais são criadas a partir da classe Toplevel
        wnd_cad_modelos = Toplevel()
        
        # define as dimensões da janela de cadastro de modelos
        wnd_cad_modelos.geometry("350x350")
        
        # define o título da janela de cadastro de modelos
        wnd_cad_modelos.title("Cadastro de Modelos")
        
        wnd_cad_modelos.modelo_label = tk.Label(wnd_cad_modelos, text="Modelo:")
        wnd_cad_modelos.modelo_label.pack()

        wnd_cad_modelos.modelo_entry = tk.Entry(wnd_cad_modelos, width=30)
        wnd_cad_modelos.modelo_entry.pack()

        wnd_cad_modelos.scrollbar = tk.Scrollbar(wnd_cad_modelos)
        #scrollbar.pack(side=RIGHT, fill=Y)
        wnd_cad_modelos.scrollbar.pack(side="right", fill="y")

        wnd_cad_modelos.modelos_listbox = tk.Listbox(wnd_cad_modelos, width=40,yscrollcommand=wnd_cad_modelos.scrollbar.set)
        wnd_cad_modelos.modelos_listbox.pack()

        wnd_cad_modelos.add_button = tk.Button(wnd_cad_modelos, text="Adicionar", command=lambda: self.add_modelo(wnd_cad_modelos))
        wnd_cad_modelos.add_button.pack()

        wnd_cad_modelos.scrollbar.config(command=wnd_cad_modelos.modelos_listbox.yview)

        self.load_modelos(wnd_cad_modelos)
        
        wnd_cad_modelos.update_button = tk.Button(wnd_cad_modelos, text="Update", command=lambda: self.update_modelo(wnd_cad_modelos))
        wnd_cad_modelos.update_button.pack()

        wnd_cad_modelos.delete_button = tk.Button(wnd_cad_modelos, text="Delete", command=lambda: self.delete_modelo(wnd_cad_modelos))
        wnd_cad_modelos.delete_button.pack()

        # coloca o foco na janela de cadastro de modelos
        wnd_cad_modelos.focus_set()
        # desabilita a janela principal enquanto a janela cadastro de modelos estiver aberta
        wnd_cad_modelos.grab_set()
        # inicia o loop de eventos da janela cadastro de empresas
        wnd_cad_modelos.mainloop()

    def iniciar(self):
        global ser
        global running
        global first
        first = True
        empresa = self.comboEmpresas.get()
        modelo = self.comboModelos.get()
        tamanho = self.tamanho_entry.get()
        numero = self.numero_entry.get()
        processo = self.processo_entry.get()

        if empresa and modelo and tamanho and numero and processo:
            global ensaio_id
            global timestampInicio
            currentDateTime = datetime.datetime.now()

            self.cursor.execute("INSERT INTO ensaios_summary (timestampInicio, empresa, modelo, tamanho, numero, processo) VALUES (?, ?, ?, ?, ?, ?)", (currentDateTime, empresa, modelo, tamanho, numero, processo))
            ensaio_id = self.cursor.lastrowid
            self.conn.commit()
            self.cursor.execute("SELECT timestampInicio FROM ensaios_summary WHERE id=?", [ensaio_id])
            ensaios = self.cursor.fetchall()
            for row in ensaios:
                timestampInicio= f"{row[0]}"
            #self.cursor.close()
            running = True
            first = True
        else:
            messagebox.showwarning("AVISO", "Por favor, preencha todos os campos.")
            running = False

        self.empresa_label.focus_set()
        ser = serial.Serial('COM3', 9600)

        def update():
            global first
            global running 
            global f_valor_lido
            global f_valor_maximo
            global f_valor_minimo

            if running:
                if ser.inWaiting() > 0:
                    linha = ser.read()
                    while ser.inWaiting() > 0:
                        linha += ser.read()
                
                    try:
                        strlinha = linha.decode("utf-8")
                        strlinha = strlinha.split('+')
                        strlinha = strlinha[1]
                        strlinha = strlinha[:-2]
                    except:
                        try:
                            strlinha = linha.decode("utf-8")
                            strlinha = strlinha.split('-')
                            strlinha = strlinha[1]
                            strlinha = strlinha[:-2]
                            strlinha = '-' + strlinha
                        except:
                            strlinha = '0'                
                    
                    try:
                        f_valor_lido = float(strlinha)
                    except:
                        f_valor_lido = 0

                    read_value.set(f_valor_lido)

                    if first:
                        f_valor_minimo = f_valor_lido
                        f_valor_maximo = f_valor_lido
                        min_value.set(f_valor_minimo)
                        max_value.set(f_valor_maximo)
                        if read_value.get() != '0':
                            first = False

                    if (f_valor_minimo > f_valor_lido):
                        f_valor_minimo = f_valor_lido
                        min_value.set(f_valor_minimo)

                    if (f_valor_maximo < f_valor_lido):
                        f_valor_maximo = f_valor_lido
                        max_value.set(f_valor_maximo)
            
                if running:
                    timer = root.after(50, update)

        if running:
            timer = root.after(50, update)

    def add_empresa(self, wnd_cad_empresas):
        empresa = wnd_cad_empresas.empresa_entry.get()
        if empresa:
            self.cursor.execute("INSERT INTO empresas (empresa) VALUES (?)", [empresa])
            self.conn.commit()
            self.load_empresas(wnd_cad_empresas)
            self.load_combo_empresas()
        else:
            messagebox.showwarning("Aviso", "Preencha o campo empresa.")

    def add_modelo(self, wnd_cad_modelos):
        modelo = wnd_cad_modelos.modelo_entry.get()
        if modelo:
            self.cursor.execute("INSERT INTO modelos (modelo) VALUES (?)", [modelo])
            self.conn.commit()
            self.load_modelos(wnd_cad_modelos)
            self.load_combo_modelos()
        else:
            messagebox.showwarning("Aviso", "Preencha o campo modelo.")

    def load_combo_ensaios(self):
        global cb_ensaios
        cb_ensaios = []
        self.cursor.execute("SELECT timestampInicio FROM ensaios_summary")
        ensaios = self.cursor.fetchall()
        for row in ensaios:
            cb_ensaios.append(f"{row[0]}")
        self.comboEnsaios['values'] = cb_ensaios
        #self.cursor.close()
    
    def load_combo_empresas(self):
        global cb_empresas
        cb_empresas = []
        self.cursor.execute("SELECT * FROM empresas")
        empresas = self.cursor.fetchall()
        for row in empresas:
            cb_empresas.append(f"{row[0]}")
        self.comboEmpresas['values'] = cb_empresas
        #self.cursor.close()

    def load_empresas(self, wnd_cad_empresas):
        wnd_cad_empresas.empresas_listbox.delete(0, tk.END)
        self.cursor.execute("SELECT * FROM empresas")
        empresas = self.cursor.fetchall()
        for row in empresas:
            wnd_cad_empresas.empresas_listbox.insert(tk.END, f"{row[0]}")
        #self.cursor.close()

    def load_combo_modelos(self):
        global cb_modelos
        cb_modelos = []
        self.cursor.execute("SELECT * FROM modelos")
        modelos = self.cursor.fetchall()
        for row in modelos:
            cb_modelos.append(f"{row[0]}")
        self.comboModelos['values'] = cb_modelos
        #self.cursor.close()
    
    def load_modelos(self, wnd_cad_modelos):
        wnd_cad_modelos.modelos_listbox.delete(0, tk.END)
        self.cursor.execute("SELECT * FROM modelos")
        modelos = self.cursor.fetchall()
        for row in modelos:
            wnd_cad_modelos.modelos_listbox.insert(tk.END, f"{row[0]}")
        #self.cursor.close()

    def load_ensaio(self):
        global cb_modelos
        self.cursor.execute("SELECT empresa, modelo, tamanho, numero, processo, valorMinimo, valorMaximo, timestampInicio FROM ensaios_summary WHERE timestampInicio=?", [self.comboEnsaios.get()])
        ensaios = self.cursor.fetchall()
        self.tamanho_entry.delete(0, tk.END)
        self.numero_entry.delete(0, tk.END)
        self.processo_entry.delete(0, tk.END)
        for row in ensaios:
            global timestampInicio
            timestampInicio=f"{row[7]}"
            self.comboEmpresas.set(f"{row[0]}")
            self.comboModelos.set(f"{row[1]}")
            self.tamanho_entry.insert(0, f"{row[2]}")
            self.numero_entry.insert(0, f"{row[3]}")
            self.processo_entry.insert(0, f"{row[4]}")
            self.menor_entry.insert(0, f"{'%.4f' % float(row[5])}")
            self.maior_entry.insert(0, f"{'%.4f' % float(row[6])}")
        #self.cursor.close()

    def parar_ensaio(self):
        #self.comboEmpresas.delete(0, tk.END)
        #self.comboModelos.delete(0, tk.END)
        #self.tamanho_entry.delete(0, tk.END)
        #self.numero_entry.delete(0, tk.END)
        #self.processo_entry.delete(0, tk.END)
        
        global ser
        global running
        global first
        ser.close()
        running = False
        first = True

        global ensaio_id
        currentDateTime = datetime.datetime.now()
        self.cursor.execute("UPDATE ensaios_summary SET timestampFim=?, valorMinimo=?, valorMaximo=? WHERE id=?", (currentDateTime, min_value.get(), max_value.get(), ensaio_id))
        self.conn.commit()
        self.load_combo_ensaios()

    def clear_empresas(self, wnd_cad_empresas):
        wnd_cad_empresas.empresa_entry.delete(0, tk.END)

    def clear_modelos(self, wnd_cad_modelos):
        wnd_cad_modelos.modelo_entry.delete(0, tk.END)

    def update_empresa(self, wnd_cad_empresas):
        selected_empresa = wnd_cad_empresas.empresas_listbox.get(tk.ACTIVE)
        if selected_empresa:
            empresa = wnd_cad_empresas.empresa_entry.get()
            if empresa:
                self.cursor.execute("UPDATE empresas SET empresa=? WHERE empresa=?", (empresa, selected_empresa))
                self.conn.commit()
                self.load_empresas(wnd_cad_empresas)
                self.clear_empresas(wnd_cad_empresas)
                self.load_combo_empresas()
            else:
                messagebox.showwarning("Aviso", "Preencha o campo empresa.")
        else:
            messagebox.showwarning("Aviso", "Selecione uma empresa.")

    def update_modelo(self, wnd_cad_modelos):
        selected_modelo = wnd_cad_modelos.modelos_listbox.get(tk.ACTIVE)
        if selected_modelo:
            modelo = self.wnd_cad_modelos.modelo_entry.get()
            if modelo:
                self.cursor.execute("UPDATE modelos SET modelo=? WHERE modelo=?", (modelo, selected_modelo))
                self.conn.commit()
                self.load_modelos(wnd_cad_modelos)
                self.clear_modelos(wnd_cad_modelos)
            else:
                messagebox.showwarning("Aviso", "Preencha o campo modelo.")
        else:
            messagebox.showwarning("Aviso", "Selecione uma modelo.")

    def delete_empresa(self, wnd_cad_empresa):
        selected_empresa = wnd_cad_empresa.empresas_listbox.get(tk.ACTIVE)
        if selected_empresa:
            self.cursor.execute("DELETE FROM empresas WHERE empresa=?", [selected_empresa])
            self.conn.commit()
            self.load_empresas(wnd_cad_empresa)
        else:
            messagebox.showwarning("Aviso", "Por favor, selecione uma empresa para deletar.")

    def delete_modelo(self, wnd_cad_modelos):
        selected_modelo = wnd_cad_modelos.modelos_listbox.get(tk.ACTIVE)
        if selected_modelo:
            self.cursor.execute("DELETE FROM modelos WHERE modelo=?", [selected_modelo])
            self.conn.commit()
            self.load_empresas(wnd_cad_modelos)
        else:
            messagebox.showwarning("Aviso", "Por favor, selecione uma empresa para deletar.")


    def delete_ensaio(self):
        selected_ensaio = self.ensaios_listbox.get(tk.ACTIVE)
        # if selected_student:
        #     student_id = int(selected_student.split(".")[0])
        #     self.cursor.execute("DELETE FROM students WHERE id=?", (student_id,))
        #     self.conn.commit()
        #     self.load_students()
        #     self.clear_entries()
        # else:
        #     messagebox.showwarning("Warning", "Please select an student to delete.")
        #running = False
    # def __del__(self):
    #     self.conn.close()

if __name__ == "__main__":
    ser = serial.Serial('COM3', 9600)
    ser.close()
    first = True
    running = False

    root = tk.Tk()
    root.geometry("500x800")
    read_value=tk.StringVar()
    max_value=tk.StringVar()
    min_value=tk.StringVar()
    cb_empresas = []
    cb_modelos = []
    cb_ensaios = []
    timestampInicio = ''
    
    app = VidrosApp(root)
    root.mainloop()

    ser.close()