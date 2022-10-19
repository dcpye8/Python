import tkinter as tk
from tkinter import ttk
import pandas as pd
import numpy as np
import pyodbc

LARGE_FONT = ("Verdana", 12)
MED_FONT = ("Verdana", 10)
PEQ_FONT = ("Verdana", 8)

class RevTingApp(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)

        #tk.Tk.iconbitmap(self, default="clienticon.ico")
        tk.Tk.wm_title(self, "Revista Tingimento")

        container = tk.Frame(self)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)  #prioridade

        self.frames = {}

        #for F in frames:

        frame = HomePage(container, self)
        self.frames[HomePage] = frame
        frame.grid(row=0, column=0, sticky="nsew")
        self.show_frame(HomePage)

    def show_frame(self, cont):

        frame = self.frames[cont]
        frame.tkraise()


class HomePage(tk.Frame):  #frame da homepage

    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        label = ttk.Label(self, text="Consulta Revista de Tingimento", font=LARGE_FONT).grid(column=1, columnspan=6)
        info_text = ttk.Label(self, text="Data no formato: ano/mês/dia", font=PEQ_FONT).grid(column=1, row=3, columnspan=2, sticky="E")
        info_text2 = ttk.Label(self, text="", font=MED_FONT).grid(column=1, row=2, columnspan=2)
        data1_text = ttk.Label(self, text="Data de:", font=MED_FONT).grid(column=1, row=4, sticky="E")
        data2_text = ttk.Label(self, text="A:", font=MED_FONT).grid(column=1, row=5, sticky="E")
        ordem_text = ttk.Label(self, text="Ordem:", font=MED_FONT).grid(column=1,row=7, sticky="E")
        artigo_text = ttk.Label(self, text="Artigo", font=MED_FONT).grid(column=1, row=6, sticky="E")
        
        row_extra = ttk.Label(self, text="").grid(column=2, row=8)
        info_text = ttk.Label(self, text="Resultados:", font=MED_FONT).grid(column=3, row=9)
        
        qtd_rev_text = ttk.Label(self, text="Qtd. Revistada (m): ", font=MED_FONT).grid(column=2, row=11, sticky="E")
        num_ord_text = ttk.Label(self, text="Núm. Ordens Revistadas: ", font=MED_FONT).grid(column=2, row=12, sticky="E")
        num_lotes_text = ttk.Label(self, text="Núm. Lotes Revistados: ", font=MED_FONT).grid(column=2, row=13, sticky="E")
        
        self.qtd_rev_text2 = ttk.Label(self, text="", font=MED_FONT)
        self.qtd_rev_text2.grid(column=3, row=11, sticky="W")

        self.num_ord_text2 = ttk.Label(self, text="")
        self.num_ord_text2.grid(column=3, row=12, sticky="W")

        self.num_lotes_text2 = ttk.Label(self, text="")
        self.num_lotes_text2.grid(column=3, row=13, sticky="W")
           
        
        for r in range(3):
            x = ttk.Label(self, text="").grid(columns=1, row=r)
            y = ttk.Label(self, text="").grid(columns=1, row=r+3)

        #caixas de input texto
        
        data1_entry=ttk.Entry(self)
        data1_entry.grid(column=2, row=4)
        
        data2_entry = ttk.Entry(self)
        data2_entry.grid(column=2, row=5)
        
        ordem_entry = ttk.Entry(self)
        ordem_entry.grid(column=2, row=7)
        
        artigo_entry = ttk.Entry(self)
        artigo_entry.grid(column=2, row=6)
            
        col_extra = ttk.Label(self, text="      ")
        col_extra.grid(column=3, row=4)


        button1 = ttk.Button(self, text="Pesquisar por Data",
                             command=lambda: self.get_data(data1_entry, data2_entry))
        button1.grid(column=3, row=4, rowspan=2)
        
        button2 = ttk.Button(self, text="Pesquisar por Ordem",
                             command=lambda: self.get_data_ordem(ordem_entry))                    
        button2.grid(column=3, row=7)

        button3 = ttk.Button(self, text="Pesquisar por Data e Artigo",
                             command=lambda: self.get_data_artigo(data1_entry, data2_entry, artigo_entry))
        button3.grid(column=3, row=6)

        
    def get_data(self, data1_entry, data2_entry):
        
        #função para filtrar o SQL server por data de inicio a fim
        data1_entry = data1_entry.get()
        data2_entry = data2_entry.get()

        bd = 'RevistaTingimento'
        
        conn = pyodbc.connect('Driver={SQL Server};'
                                        'Server=srvsql02;'
                                        'Database=' + bd + ';'
                                        'Trusted_Connection=yes;')
        

        sql_info = "SELECT Numero, Date, Machine, Utilisateur, Etat, Metrage, Choix, I4, I5, I6, I7, D1, Metrage1, Longueur, Bonus1, DP, Libelle, Article, NumeroOF, Obs FROM ViewCoupe01 WHERE Date >= '"+str(data1_entry)+"' AND Date <= '"+str(data2_entry)+"'"
                    
        df = pd.read_sql(sql_info, conn)   
        conn.close()

        df = pd.DataFrame(df)
        
        df.columns = ['Lote_orig', 'Data', 'Maquina', 'Utilizador', 'Estado', 'Metragem', 'Escolha', 'Larg_teorica',
                      'Larg_real', 'Larg_real2', 'Larg_real3', 'Metragem2', 'Metragem_def', 'Comprimento_def', 'Bonus',
                      'DP', 'Defeito', 'Artigo', 'Ordem', 'Obs']

        # return da soma do revistado na data escolhida
        df['lotes_dupl'] = df.duplicated(keep="last", subset="Lote_orig")
        df['ord_dupl'] = df.duplicated(keep="last", subset="Ordem")
        
        qtd_rev = df[(df.lotes_dupl == False)].sum()['Metragem']
        num_lotes = df[(df.lotes_dupl == False)].count()['Metragem']
        num_ord = df[(df.ord_dupl == False)].count()['Metragem']
        
        df['n_defeitos'] = 1

        rev_ting = pd.pivot_table(df, values=['n_defeitos', 'Comprimento_def'],
                                  index=['Ordem', 'Artigo', 'Lote_orig', 'Metragem', 'Data', 'Obs', 'Larg_teorica', 'Larg_real', 'Defeito',
                                         'Escolha'], fill_value=None, aggfunc=np.sum).astype(int)
        rev_ting.sort_values(by=['Ordem', 'Lote_orig', 'n_defeitos'], inplace=True, ascending=False)
        rev_ting.to_excel("Rel_RevTing.xlsx")
        
        self.qtd_rev_text2.config(text=str(int(qtd_rev))+" m")
        self.num_ord_text2.config(text=str(int(num_ord))+" ordens")
        self.num_lotes_text2.config(text=str(int(num_lotes))+" lotes")
        
    def get_data_ordem(self, ordem_entry):

        #função para filtrar o SQL server por ordem
        
        ordem_entry = ordem_entry.get()
        ordem_entry = "'"+ordem_entry+"'"
        bd = 'RevistaTingimento'
        
        conn = pyodbc.connect('Driver={SQL Server};'
                                            'Server=srvsql02;'
                                            'Database=' + bd + ';'
                                            'Trusted_Connection=yes;')
        

        sql_info = "SELECT Numero, Date, Machine, Utilisateur, Etat, Metrage, Choix, I4, I5, I6, I7, D1, Metrage1, Longueur, Bonus1, DP, Libelle, Article, NumeroOF, Obs FROM ViewCoupe01 WHERE NumeroOf="+str(ordem_entry)
            
        df = pd.read_sql(sql_info, conn)   
            
        conn.close()

        df = pd.DataFrame(df)
        df.columns = ['Lote_orig', 'Data', 'Maquina', 'Utilizador', 'Estado', 'Metragem', 'Escolha', 'Larg_teorica',
                      'Larg_real', 'Larg_real2', 'Larg_real3', 'Metragem2', 'Metragem_def', 'Comprimento_def', 'Bonus',
                      'DP', 'Defeito', 'Artigo', 'Ordem', 'Obs']

        df['lotes_dupl'] = df.duplicated(keep="last", subset="Lote_orig")
        df['ord_dupl'] = df.duplicated(keep="last", subset="Ordem")
        
        qtd_rev = df[(df.lotes_dupl == False)].sum()['Metragem']
        num_lotes = df[(df.lotes_dupl == False)].count()['Metragem']
        num_ord = df[(df.ord_dupl == False)].count()['Metragem']
        
        df['n_defeitos'] = 1
        
        rev_ting = pd.pivot_table(df, values=['n_defeitos', 'Comprimento_def'],
                                  index=['Ordem', 'Artigo', 'Lote_orig', 'Metragem', 'Data','Obs', 'Larg_teorica', 'Larg_real', 'Defeito',
                                         'Escolha'], fill_value=None, aggfunc=np.sum).astype(int)
        rev_ting.sort_values(by=['Ordem', 'Lote_orig', 'n_defeitos'], inplace=True, ascending=False)
        rev_ting.to_excel("Rel_RevTing "+ordem_entry+".xlsx")

        self.qtd_rev_text2.config(text=str(int(qtd_rev))+" m")
        self.num_ord_text2.config(text=str(int(num_ord))+" ordens")
        self.num_lotes_text2.config(text=str(int(num_lotes))+" lotes")

    def get_data_artigo(self, data1_entry, data2_entry, artigo_entry):

    #função para filtrar o SQL server por data de inicio a fim e artigo
        
        data1_entry = data1_entry.get()
        data2_entry = data2_entry.get()
        artigo_entry = artigo_entry.get()
        artigo_entry = "'"+artigo_entry+"'"
        
        bd = 'RevistaTingimento'
        
        conn = pyodbc.connect('Driver={SQL Server};'
                                        'Server=srvsql02;'
                                        'Database=' + bd + ';'
                                        'Trusted_Connection=yes;')
        

        sql_info = "SELECT Numero, Date, Machine, Utilisateur, Etat, Metrage, Choix, I4, I5, I6, I7, D1, Metrage1, Longueur, Bonus1, DP, Libelle, Article, NumeroOF, Obs FROM ViewCoupe01 WHERE Article =" +str(artigo_entry)+" AND Date >= '"+str(data1_entry)+"' AND Date <= '"+str(data2_entry)+"'"
                    
        df = pd.read_sql(sql_info, conn)   
        conn.close()

        df = pd.DataFrame(df)
        df.columns = ['Lote_orig', 'Data', 'Maquina', 'Utilizador', 'Estado', 'Metragem', 'Escolha', 'Larg_teorica',
                      'Larg_real', 'Larg_real2', 'Larg_real3', 'Metragem2', 'Metragem_def', 'Comprimento_def', 'Bonus',
                      'DP', 'Defeito', 'Artigo', 'Ordem', 'Obs']
        
        df['lotes_dupl'] = df.duplicated(keep="last", subset="Lote_orig")
        df['ord_dupl'] = df.duplicated(keep="last", subset="Ordem")
        
        qtd_rev = df[(df.lotes_dupl == False)].sum()['Metragem']
        num_lotes = df[(df.lotes_dupl == False)].count()['Metragem']
        num_ord = df[(df.ord_dupl == False)].count()['Metragem']
        
        df['n_defeitos'] = 1

        rev_ting = pd.pivot_table(df, values=['n_defeitos', 'Comprimento_def'],
                                  index=['Ordem', 'Artigo', 'Lote_orig', 'Metragem', 'Data', 'Obs', 'Larg_teorica', 'Larg_real', 'Defeito',
                                         'Escolha'], fill_value=None, aggfunc=np.sum).astype(int)
        rev_ting.sort_values(by=['Ordem', 'Lote_orig', 'n_defeitos'], inplace=True, ascending=False)
        rev_ting.to_excel("Rel_RevTing "+artigo_entry+".xlsx")

        self.qtd_rev_text2.config(text=str(int(qtd_rev))+" m")
        self.num_ord_text2.config(text=str(int(num_ord))+" ordens")
        self.num_lotes_text2.config(text=str(int(num_lotes))+" lotes")

if __name__ == "__main__":

    app = RevTingApp()
    app.mainloop()
