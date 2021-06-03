import firebirdsql
import csv
import tkinter as tk
from tkinter import *

#Estudo do Banco de Dados Medmanager
#Nivea quer: Data(ok), Nome(ok),Convênio(ok), Exames(ok), Médico Solicitante(ok), Médico que realizou o exame(ok), Observações(ok)
# AGD_AINC(0) - id, AGD_DATA(1) - data do exame, AGD_PACT(2) - id do paciente ,AGD_FICH(3) - , AGD_NOME(4) - nome do paciente, AGD_NSPQ(5) - nome padronizado, AGD_TELF(6) - telefone, AGD_CONV(7) - convênios, AGD_CMED(8) - id médica que realizou, AGD_PRC1(9) - procedimento 1 , AGD_PRC2(10) - procedimento 2, AGD_PRC3(11) - procedimento 3, AGD_PRC4(12) - procedimento 4, AGD_PRC5(13) - procedimento 5, AGD_PRC6(14) - procedimento 6, AGD_OBSR(15) - observarção
#Vai usar as tabelas: convenios,atnlaudos
#talvez pacienteconv tenha os numeros das carteiras
	
def start(data_ini, data_fim):
	conn = firebirdsql.connect(user="SYSDBA", password="masterkey",database="C:\\Users\\victo\\Documents\\Database\\db2\\Database\\MEDMANAG.FDB", host="localhost")
	cur = conn.cursor()
	print(data_ini, data_fim)
	consulta = f"""SELECT ATNDLAUDOS.LDO_DATA, PACIENTES.PAC_NPSQ, CONVENIOS.CNV_DESC, ATNDLAUDOS.LDO_MATR, PROCEDIMENT.PRC_DESC, MEDSOLICIT.MED_NPSQ, ATNDLAUDOS.LDO_CMED FROM ATNDLAUDOS 
				LEFT JOIN PACIENTES ON ATNDLAUDOS.LDO_PACT = PACIENTES.PAC_AINC
				LEFT JOIN CONVENIOS ON ATNDLAUDOS.LDO_CONV = CONVENIOS.CNV_AINC
				LEFT JOIN PROCEDIMENT ON ATNDLAUDOS.LDO_PROC = PROCEDIMENT.PRC_AINC
				LEFT JOIN MEDSOLICIT ON ATNDLAUDOS.LDO_CLIN = MEDSOLICIT.MED_AINC
				WHERE LDO_DATA BETWEEN '{data_ini}' AND '{data_fim}' 
				ORDER BY LDO_DATA"""
	cur.execute(consulta)
	with open(f"Pacientes_{data_ini}--{data_fim}.csv","w") as csv_file:
		csv_writer = csv.writer(csv_file)
		csv_writer.writerow(["Data","Nome do Paciente", "Convênio", "Numero da Guia", "Exame Realizado", "Médico Solicitante", "Médica Que Atendeu"])
		for c in cur.fetchall():
			if(c[5]==1):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Gerusa - Ultrasom",c[6]])
			elif(c[5]==2):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Gerusa - Mamografia",c[6]])
			elif(c[5]==3):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Valéria - Ultrasom",c[6]])
			elif(c[5]==4):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Valéria - Mamografia",c[6]])
			elif(c[5]==5):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Procedimentos",c[6]])
			elif(c[5]==6):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Laurise - Ultrasom",c[6]])
			elif(c[5]==7):
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],"Laurise - Mamografia",c[6]])
			else:
				csv_writer.writerow([c[0],c[1],c[2],c[3],c[4],c[5],c[6]])
		conn.close()

	print("Conexão concluida")
	root.destroy()



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







#LEFT JOIN AGDIARIA ON ATNDLAUDOS.LDO_PACT = AGDIARIA.AGD_PACT AND ATNDLAUDOS.LDO_DATA = CAST(AGDIARIA.AGD_DATA AS DATE)