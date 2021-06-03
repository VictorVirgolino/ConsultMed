import firebirdsql
import tkinter as tk
from tkinter import *
import xlsxwriter
from datetime import datetime

# Helper criação do medmanager valeria, geruza, laurise, procedimentos para uso no balanceUS


def start(data_ini, data_fim):
	conn = firebirdsql.connect(user="SYSDBA", password="masterkey",database="C:\\Users\\victo\\Documents\\Database\\db2\\Database\\MEDMANAG.FDB", host="localhost")
	cur = conn.cursor()
	print(data_ini, data_fim)
	consulta = f"""SELECT ATNDLAUDOS.LDO_DATA, PACIENTES.PAC_NPSQ, PROCEDIMENT.PRC_DESC, ATNDLAUDOS.LDO_CMED FROM ATNDLAUDOS 
				LEFT JOIN PACIENTES ON ATNDLAUDOS.LDO_PACT = PACIENTES.PAC_AINC
				LEFT JOIN PROCEDIMENT ON ATNDLAUDOS.LDO_PROC = PROCEDIMENT.PRC_AINC
				WHERE LDO_DATA BETWEEN '{data_ini}' AND '{data_fim}' 
				ORDER BY LDO_DATA"""
	cur.execute(consulta)

	wkValeria = xlsxwriter.Workbook(f"medmanager valeria {data_ini}--{data_fim}.xlsx")
	wkGerusa = xlsxwriter.Workbook(f"medmanager gerusa {data_ini}--{data_fim}.xlsx")
	wkLaurise = xlsxwriter.Workbook(f"medmanager laurise {data_ini}--{data_fim}.xlsx")
	wkProcedimentos = xlsxwriter.Workbook(f"medmanager procedimentos {data_ini}--{data_fim}.xlsx")
	wsV = wkValeria.add_worksheet()
	wsG = wkGerusa.add_worksheet()
	wsL = wkLaurise.add_worksheet()
	wsP = wkProcedimentos.add_worksheet()
	formatG = wkGerusa.add_format({'num_format': 'dd/mm/yy'})
	formatV = wkValeria.add_format({'num_format': 'dd/mm/yy'})
	formatL = wkLaurise.add_format({'num_format': 'dd/mm/yy'})
	formatP = wkProcedimentos.add_format({'num_format': 'dd/mm/yy'})

	lista = ["Data", "Nome do Paciente", "Procedimento", "Valor"]
	col = 0
	sheet = wsV
	rowV = 0
	rowG = 0
	rowL = 0
	rowP = 0
	
	for coluna in lista:
		wsV.write(0, col, coluna)
		wsG.write(0, col, coluna)
		wsL.write(0, col, coluna)
		wsP.write(0, col, coluna)
		col += 1

	for c in cur.fetchall():
		if (c[3] == 1 or c[3]== 2):
			wsG.write(rowG, 0, c[0], formatG)
			for x in range(1,3):
				wsG.write(rowG, x, c[x])
			rowG += 1
		elif (c[3] == 3 or c[3]== 4):
			wsV.write(rowV, 0, c[0], formatV)
			for y in range(1,3):
				wsV.write(rowV, y, c[y])
			rowV += 1
		elif (c[3] == 5):
			wsP.write(rowP, 0, c[0], formatP)
			for w in range(1,3):
				wsP.write(rowP, w, c[w])
			rowP += 1
		elif (c[3] == 6 or c[3]== 7):
			wsL.write(rowL, 0, c[0], formatL)
			for z in range(1,3):
				wsL.write(rowL, z, c[z])
			rowL += 1
		
		
	
	wkValeria.close()
	wkGerusa.close()
	wkLaurise.close()
	wkProcedimentos.close()
	
	print("Done")



#Tkinter
root = tk.Tk()
root.title("ConsultMed")
root.geometry("300x200")
root.configure(background="#dde")

lb1 =Label(root, text="Data Inicial: ",background="#dde", foreground="#009",anchor=W)
data_ini = Entry(root)
lb2 =Label(root, text="Data Final: ",background="#dde", foreground="#009",anchor=W)
data_fim = Entry(root)
bt =Button(root, text="Start", command=lambda: start(data_ini.get(), data_fim.get()), bg="#AFFFB7", fg='#8421FC')
lb1.pack()
data_ini.pack()
lb2.pack()
data_fim.pack()
bt.pack()
root.mainloop()





"""
wk = xlsxwriter.Workbook("Teste.xlsx")
ws = wk.add_worksheet()

expenses = (
    ['Rent', 1000],
    ['Gas',   100],
    ['Food',  300],
    ['Gym',    50],
)

row = 0
col = 0

for item, cost in expenses:
	ws.write(row, col, item)
	ws.write(row, col+1, cost)
	row += 1

wk.close()
"""