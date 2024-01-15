from tkinter import *
from tkcalendar import DateEntry
from tkinter import ttk
import openpyxl
import os 
import random


root = Tk()


class Aplication():
    book = openpyxl.Workbook()
    template_name = "Template.xlsx"
    file_workbook = "Banco de Dados"
    template_wb = openpyxl.load_workbook(template_name)
    template_ws = template_wb[file_workbook]
    line = 1
    months = {
        1: "Janeiro",
        2: "Fevereiro",
        3: "Março",
        4: "Abril",
        5: "Maio",
        6: "Junho",
        7: "Julho",
        8: "Agosto",
        9: "Setembro",
        10: "Outubro",
        11: "Novembro",
        12: "Dezembro",
        }

    def find_table(self, filename):
        self.wb = openpyxl.load_workbook(filename)
        self.ws = self.wb[self.file_workbook]
        last_row = self.ws.max_row
        self.line = last_row + 1
    
    def new_table(self, data):
        for i in range(1, 11):
            self.template_ws.cell(self.line, i, data[i - 1])
        self.line = self.line + 1    
    def append_table(self, data):
        for i in range(1, 11):
            self.ws.cell(self.line, i, data[i - 1])
        self.line = self.line + 1
    def check_table(self, filename):
        return os.path.isfile(filename)   
    def clean_table(self):
        self.entry_id.delete(0, END)
        self.entry_data.delete(0, END)
        self.entry_copias_br.delete(0, END)
        self.entry_copias_r.delete(0, END)
        self.entry_perdas_br.delete(0, END)
        self.entry_perdas_r.delete(0, END)
        self.entry_pg_dinheiro.delete(0, END)
        self.entry_pg_pix.delete(0, END)
        self.entry_cptotal_br.delete(0, END)
        self.entry_cptotal_r.delete(0, END)
        self.entry_id.delete(0, END)
    def insert_tree(self, sheet):
        rows = sheet.iter_rows()
        self.list_print.delete(*self.list_print.get_children())
        for row in rows:
            values = [cell.value for cell in row]
            self.list_print.insert('', 'end', values=values)
    def on_double_click(self, event):
        index = event.widget.selection()[0]
        item = event.widget.item(index)
        data = item['values']
        self.clean_table()
        self.insert_entry(data)
    def id_generator(self, tamanho):
        caracteres = "abcdefghijklmnopqrstuvwxyz0123456789ABCDEFGHIJKLMNOPQRSTUVXYZ"
        return "".join(random.choice(caracteres) for _ in range(tamanho))
    def filter_column(self):
        string = self.entry_id.get()
        column = self.ws['A']
        for cell in column:
            row_value = cell.value
            if string in row_value:
                row_cells = self.ws[cell.row]
                row_data = [cell.value for cell in row_cells]
        self.clean_table()
        self.insert_entry(row_data)
    def insert_entry(self, array):
        self.entry_id.insert(0, array[0])
        self.entry_data.insert(0, array[1])
        self.entry_copias_br.insert(0, array[2])
        self.entry_copias_r.insert(0, array[3])
        self.entry_perdas_br.insert(0, array[4])
        self.entry_perdas_r.insert(0, array[5])
        self.entry_cptotal_br.insert(0, array[6])
        self.entry_cptotal_r.insert(0, array[7])
        # self.entry_pg_dinheiro.insert(0, array[8])
        # self.entry_pg_pix.insert(0, array[9])




    def save_table(self):
        month = str(self.months[int(self.entry_data.get().split("/")[0])])
        year = str(self.entry_data.get().split("/")[2])
        name = month + "20" + year + ".xlsx"
        id = self.id_generator(4)
        data = []
        data.append(id)
        data.append(self.entry_data.get())
        data.append(int(self.entry_copias_br.get()))
        data.append(int(self.entry_copias_r.get()))
        data.append(int(self.entry_perdas_br.get()))
        data.append(int(self.entry_perdas_r.get()))
        data.append(int(self.entry_cptotal_br.get()))
        data.append(int(self.entry_cptotal_r.get()))
        data.append("R$ " + str(float(self.entry_pg_dinheiro.get())))
        data.append("R$ " + str(float(self.entry_pg_pix.get())))

        if not self.check_table(name):
            print("Cria a planilha")
            self.line = 1
            self.new_table(data)
            self.template_wb.save(name) ; print("salvo")
            self.insert_tree(self.template_ws)
            print(name)
        else:
            print("A planilha já existe")
            self.find_table(name)
            self.append_table(data) 
            self.wb.save(name) ; print("salvo")
            self.insert_tree(self.ws)
        self.clean_table()
        
        


class Window(Aplication):

    def __init__(self):
        self.create_window()
        self.create_frames()
        self.entry_label()
        self.output_list()
        self.buttons()

    def create_window(self):
        root.title("Impressão Digital")
        root.geometry("788x588")
        root.resizable(False, False)
        root.configure(background= "#700316")
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

    def create_frames(self):
        self.frame_1 = Frame(root, bg="white", width="780", height="290", highlightbackground="#70032c", highlightthickness=6)
        self.frame_1.grid(column=0, row=0, sticky="n")
        self.frame_2 = Frame(root, bg="white", width="780", height="290", highlightbackground="#70032c", highlightthickness=6)
        self.frame_2.grid(column=0, row=1, sticky="s")

    def buttons(self):

        self.bt_buscar = Button(self.frame_1, text="Buscar", command=self.filter_column)
        self.bt_buscar.place(relx=0.624, rely=0.1, relheight=0.1, relwidth=0.08)

        self.bt_salvar = Button(self.frame_1, text="Salvar", command=self.save_table)
        self.bt_salvar.place(relx=0.015, rely=0.4, relheight=0.1, relwidth=0.1)

        self.bt_apagar = Button(self.frame_1, text="Apagar")
        self.bt_apagar.place(relx=0.015, rely=0.55, relheight=0.1, relwidth=0.1)

        self.bt_atualizar = Button(self.frame_1, text="Atualizar", )
        self.bt_atualizar.place(relx=0.015, rely=0.865, relheight=0.1, relwidth=0.1)

        self.bt_limpar = Button(self.frame_1, text="Limpar", command=self.clean_table)
        self.bt_limpar.place(relx=0.885, rely=0.865, relheight=0.1, relwidth=0.1)

    def entry_label(self):
        self.label_data = Label(self.frame_1, text="Data")
        self.label_data.place(relx=0.135, rely=0.25, relheight=0.1, relwidth=0.1)
        self.entry_data = DateEntry(self.frame_1, selectmode="day")
        self.entry_data.place(relx=0.135, rely=0.4, relheight=0.1, relwidth=0.1)

        self.label_id = Label(self.frame_1, text="ID")
        self.label_id.place(relx=0.714, rely=0.1, relheight=0.1, relwidth=0.08)
        self.entry_id = Entry(self.frame_1)
        self.entry_id.place(relx=0.811, rely=0.1, relheight=0.1, relwidth=0.1)


        self.label_copias = Label(self.frame_1, text="Cópias")
        self.label_copias.place(relx=0.25, rely=0.25, relheight=0.1, relwidth=0.17)
        self.label_copias_br = Label(self.frame_1, text="Brother")
        self.label_copias_br.place(relx=0.25, rely=0.4, relheight=0.1, relwidth=0.08)
        self.label_copias_r = Label(self.frame_1, text="Ricoh")
        self.label_copias_r.place(relx=0.34, rely=0.4, relheight=0.1, relwidth=0.08)
        self.entry_copias_br = Entry(self.frame_1)
        self.entry_copias_br.place(relx=0.25, rely=0.55, relheight=0.1, relwidth=0.08)
        self.entry_copias_r = Entry(self.frame_1)
        self.entry_copias_r.place(relx=0.34, rely=0.55, relheight=0.1, relwidth=0.08)

        self.label_perdas = Label(self.frame_1, text="Perdas")
        self.label_perdas.place(relx=0.437, rely=0.25, relheight=0.1, relwidth=0.17)
        self.label_perdas_br = Label(self.frame_1, text="Brother")
        self.label_perdas_br.place(relx=0.437, rely=0.4, relheight=0.1, relwidth=0.08)
        self.label_perdas_r = Label(self.frame_1, text="Ricoh")
        self.label_perdas_r.place(relx=0.527, rely=0.4, relheight=0.1, relwidth=0.08)
        self.entry_perdas_br = Entry(self.frame_1)
        self.entry_perdas_br.place(relx=0.437, rely=0.55, relheight=0.1, relwidth=0.08)
        self.entry_perdas_r = Entry(self.frame_1)
        self.entry_perdas_r.place(relx=0.527, rely=0.55, relheight=0.1, relwidth=0.08)

        self.label_cptotal = Label(self.frame_1, text="Cópias totais")
        self.label_cptotal.place(relx=0.624, rely=0.25, relheight=0.1, relwidth=0.17)
        self.label_cptotal_br = Label(self.frame_1, text="Brother")
        self.label_cptotal_br.place(relx=0.624, rely=0.4, relheight=0.1, relwidth=0.08)
        self.label_cptotal_r = Label(self.frame_1, text="Ricoh")
        self.label_cptotal_r.place(relx=0.714, rely=0.4, relheight=0.1, relwidth=0.08)
        self.entry_cptotal_br = Entry(self.frame_1)
        self.entry_cptotal_br.place(relx=0.624, rely=0.55, relheight=0.1, relwidth=0.08)
        self.entry_cptotal_r = Entry(self.frame_1)
        self.entry_cptotal_r.place(relx=0.714, rely=0.55, relheight=0.1, relwidth=0.08)


        self.label_pagamento = Label(self.frame_1, text="Pagamento")
        self.label_pagamento.place(relx=0.811, rely=0.25, relheight=0.1, relwidth=0.17)
        self.label_pg_dinheiro = Label(self.frame_1, text="Dinheiro")
        self.label_pg_dinheiro.place(relx=0.811, rely=0.4, relheight=0.1, relwidth=0.08)
        self.label_pg_pix = Label(self.frame_1, text="Pix")
        self.label_pg_pix.place(relx=0.9, rely=0.4, relheight=0.1, relwidth=0.08)
        self.entry_pg_dinheiro = Entry(self.frame_1)
        self.entry_pg_dinheiro.place(relx=0.811, rely=0.55, relheight=0.1, relwidth=0.08)
        self.entry_pg_pix = Entry(self.frame_1)
        self.entry_pg_pix.place(relx=0.9, rely=0.55, relheight=0.1, relwidth=0.08)

    def output_list(self):
        self.list_print = ttk.Treeview(self.frame_2, height=5, columns=("col1","col2","col3","col4","col5","col6","col7","col8","col9","col10"))
        self.list_print.place(relx=0.01, rely=0.1, relwidth=0.95, relheight=0.85)
        
        self.list_scroll = Scrollbar(self.frame_2, orient="vertical")
        self.list_print.configure(yscroll=self.list_scroll.set)
        self.list_scroll.place(relx=0.96,rely=0.1,relheight=0.85,relwidth=0.035)

        self.list_print.heading("#0", text="")
        self.list_print.heading("#1", text="ID")
        self.list_print.heading("#2", text="Data")
        self.list_print.heading("#3", text="Brother")
        self.list_print.heading("#4", text="Ricoh")
        self.list_print.heading("#5", text="Brother")
        self.list_print.heading("#6", text="Ricoh")
        self.list_print.heading("#7", text="Brother")
        self.list_print.heading("#8", text="Ricoh")
        self.list_print.heading("#9", text="Pix")
        self.list_print.heading("#10", text="Dinheiro")

        self.list_print.column("#0", width=1)
        self.list_print.column("#1", width=50)
        self.list_print.column("#2", width=80)
        self.list_print.column("#3", width=60)
        self.list_print.column("#4", width=60)
        self.list_print.column("#5", width=60)
        self.list_print.column("#6", width=60)
        self.list_print.column("#7", width=60)
        self.list_print.column("#8", width=60)
        self.list_print.column("#9", width=60)
        self.list_print.column("#10", width=60)

        self.list_print.bind('<Double-Button-1>', self.on_double_click)

        self.frame2_label_copias = Label(self.frame_2, text="Copias", borderwidth=2, relief="solid")
        self.frame2_label_copias.place(relx=0.232, rely=0.001, relheight=0.1 ,relwidth=0.182)
        self.frame2_label_perdas = Label(self.frame_2, text="Perdas", borderwidth=2, relief="solid")
        self.frame2_label_perdas.place(relx=0.414, rely=0.001, relheight=0.1 ,relwidth=0.182)
        self.frame2_label_totais = Label(self.frame_2, text="Totais", borderwidth=2, relief="solid")
        self.frame2_label_totais.place(relx=0.597, rely=0.001, relheight=0.1 ,relwidth=0.18)
        self.frame2_label_dinheiro = Label(self.frame_2, text="Dinheiro", borderwidth=2, relief="solid")
        self.frame2_label_dinheiro.place(relx=0.777, rely=0.001, relheight=0.1 ,relwidth=0.182)


def main():
    Window()

    root.mainloop()

if __name__ == "__main__":
    main()