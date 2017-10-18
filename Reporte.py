# !/usr/bin/env python3
# -*- enconding: utf-8 -*-

import os
import datetime
import time
import xlsxwriter
import calendar
import operator

from collections import OrderedDict
from tabulate import tabulate
from peewee import *
from tqdm import *

hoy = datetime.datetime.combine(datetime.datetime.today(),datetime.time())
inicio = datetime.datetime(hoy.year, 1, 1, 0, 0, 0, 0)

ano = hoy.year
mes = hoy.month
calendariojuliano = []
if calendar.isleap(ano):
	for x in range(1,367):
		calendariojuliano.append(str(inicio.strftime("%W"))+"-"+str(x).rjust(3, "0")+"-"+str(inicio.strftime("%d%m%Y")))
		inicio = inicio + datetime.timedelta(1)
else:
	for x in range(1,366):
		calendariojuliano.append(str(inicio.strftime("%W"))+"-"+str(x).rjust(3, "0")+"-"+str(inicio.strftime("%d%m%Y")))
		inicio = inicio + datetime.timedelta(1)
inicio = datetime.datetime(hoy.year, 1, 1, 0, 0, 0, 0)

semana_anterior = int(hoy.strftime("%W")) -1
semana_anterior = str(semana_anterior)

vd = 0
for x in calendariojuliano:
	if x[0:2] == semana_anterior:
		if vd == 0:
			fechareporteinicial = x[7:]
		vd += 1
		if vd == 7:
			fechareportefinal = x[7:]


app_dir = os.path.dirname(os.path.realpath(__file__))
semana = datetime.datetime.now().strftime("%W")
fecharh = datetime.datetime.now().strftime("%d%m%y%H%M%S")


NombreDB = 'Reporte_OperacionS' + semana_anterior + '.db'
db = SqliteDatabase(NombreDB)


class Agencia_Aduanal(Model):
	""" Creacion de base de datos de operaciones del agente aduanal """
	numeropedimento = IntegerField()
	movimiento = CharField()
	clavedocumento = CharField()
	rfc = CharField()
	archivom = CharField()

	class Meta:
		database = db

class Aduana(Model):
	""" Creacion de base de datos operaciones de aduana """
	npedimento_validado = ForeignKeyField(Agencia_Aduanal, related_name='patentesvalidados')
	acuse = CharField()
	archivof = CharField()

	class Meta:
		database = db

class Banco(Model):
	""" Creacion de base de datos operaciones de banco """
	npedimento_pagado = ForeignKeyField(Agencia_Aduanal, related_name='patentespaga')
	patenteb = IntegerField()
	firmabancaria = CharField()
	fechapago = CharField()
	archivop = CharField()

	class Meta:
		database = db

def leer_archivos(diajuliano):
	"""Leer un dia juliano"""
	diajuliano = str(diajuliano)
	print("Proceso Lectura de archivos " + diajuliano + " espere por favor ... \n")
	ciclo = str(datetime.date.today().year) 
	if len(diajuliano) == 1 or len(diajuliano) == 2:
		diajuliano = diajuliano.rjust(3, "0")
	carpeta = "Caaarem_pre/Concentra/Dia" + ciclo[2:4] + str(diajuliano) + "/Aduana51"
	if os.path.isdir(carpeta):
		lista = os.listdir(carpeta)
		for x in tqdm(range(len(lista))):
			archivo = lista[x]
			if archivo[0].upper() == "M":
				leer = open(carpeta + "/" + archivo, "r+", encoding='iso-8859-1')
				linea = leer.readline()
				while linea != "":
					try:
						if linea[0:3] == "501":
							if int(linea[21:22]) == 1: 
								movival = "Impo"
							elif int(linea[21:22]) == 2:
								movival = "Expo"
							Agencia_Aduanal.create(
								numeropedimento = linea[4:8]+linea[9:16], 
								movimiento = movival, 
								clavedocumento = linea[23:25], 
								rfc = linea[31:43], 
								archivom = archivo)
						linea = leer.readline()
					except ValueError:
						print("Archivo tiene un error " + archivo)
						raise
				leer.close()
			elif archivo[0].upper() == "F":
				leer = open(carpeta + "/" + archivo, "r+")
				linea = leer.readline()
				while linea != "":
					if linea[0] == "F":
						Aduana.create(
							npedimento_validado = archivo[1:5]+linea[1:8],
							acuse = linea[8:16],
							archivof = archivo)
					linea = leer.readline()
				leer.close()
			elif archivo[0].upper() == "A":
				leer = open(carpeta + "/" + archivo, "r+")
				linea = leer.readline()
				while linea != "":
					if linea[0:2] == "30":
						Banco.create(
							npedimento_pagado = linea[4:8]+linea[8:15],
							patenteb =  linea[4:8],
							firmabancaria = linea[40:50],
							fechapago = linea[50:58],
							archivop = archivo)
					linea = leer.readline()
				leer.close()
	print("\n")

def creacion_conexion():
	db.connect()
	db.create_tables([Agencia_Aduanal,Aduana,Banco],safe=True)

def menu_loop():
	opcion = None
	while opcion != 'q':
		print()
		print("Presione la tecla q para salir")
		for key,value in menu.items():
			print("{} - {}".format(key, value.__doc__))
		opcion = input("Seleccione una opcion: ").lower().strip()
		print()
		if opcion in menu:
		 	menu[opcion]()


def leer_variosj():
	"""Lee varios dias julianos"""
	ano = hoy.year
	mes = hoy.month
	print("Fecha del Inicio")
	ano = int(input("Año del reporte :"))
	mes = int(input("Mes del reporte :"))
	dia = int(input("Día del reporte :"))
	inicioj = datetime.datetime(ano, mes, dia)
	print("Fecha del Final")
	ano = int(input("Año del reporte :"))
	mes = int(input("Mes del reporte :"))
	dia = int(input("Día del reporte :"))
	finj = datetime.datetime(ano, mes, dia)
	inicioj = int(inicioj.strftime("%j"))
	finj = int(finj.strftime("%j"))
	for x in range(inicioj,finj + 1):
		leer_archivos(x)



def buscar_pedimento():
	"""Busca un pedimento"""
	encabezado = ["NumPedim","Movi","TP","CvD","RFC","Ac Ele","Archivo F","FecPag","Archivo B"]
	print("Favor de teclaer")
	patente = input("Numero de la pente : ")
	consecutivo = input("Numero de pedimento : ")
	pedimentos = (Agencia_Aduanal
			.select(Agencia_Aduanal,Aduana,Banco)
			.join(Aduana, on=(Agencia_Aduanal.numeropedimento==Aduana.npedimento_validado))
			.switch(Agencia_Aduanal)
			.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
			.where(Agencia_Aduanal.numeropedimento == patente + consecutivo))
	tabla_pedimentos= []
	for pedimento in pedimentos:
		registro = [pedimento.numeropedimento, 
					pedimento.movimiento, 
					pedimento.clavedocumento, 
					pedimento.rfc, 
					pedimento.aduana.acuse,
					pedimento.aduana.archivof,
					pedimento.banco.fechapago,
					pedimento.banco.archivop]
		tabla_pedimentos.append(registro)
	print(tabulate(tabla_pedimentos, encabezado, tablefmt='fancy_grid'))


def reportepatente():
	"""Listado de todos los pedimentos generados"""
	encabezado = ["NuPedim", "TP", "CvD","RFC","Archivo M"]
	tabla_pedimentos = []
	for pedimento in Agencia_Aduanal.select().paginate(1,12):
		registro = [pedimento.numeropedimento,
			pedimento.movimiento,
			pedimento.clavedocumento,
			pedimento.rfc, 
			pedimento.archivom]
		tabla_pedimentos.append(registro)
	print(tabulate(tabla_pedimentos, encabezado, tablefmt='fancy_grid'))

def reportepedimentovalidado():
	"""Listado de pedimentos validados"""
	pedimentos = (Agencia_Aduanal
					.select(Agencia_Aduanal, Aduana)
					.join(Aduana, on=(Agencia_Aduanal.numeropedimento==Aduana.npedimento_validado))
					.paginate(3, 20)
					)
	encabezado= ["NumPedim", "Mov", "TP", "CvD", "RFC" "Archivo M", "Acuse Elc", "Archivo Res"]
	tabla_pedimentos= []
	for pedimento in pedimentos:
		registro = [pedimento.numeropedimento, 
					pedimento.movimiento, 
					pedimento.clavedocumento, 
					pedimento.rfc, 
					pedimento.archivom,
					pedimento.aduana.acuse,
					pedimento.aduana.archivof]
		tabla_pedimentos.append(registro)
	print(tabulate(tabla_pedimentos, encabezado, tablefmt='fancy_grid'))

def reportepedimentopagado():
	"""Listado de pedimentos pagados"""
	pedimentospagados = (Agencia_Aduanal
					.select(Agencia_Aduanal, Banco)
					.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
					)
	print("NumeroPedim - TP - CD - Re.Fed.Con. - Archivo GenM - Firma Banc - Archivo Res")
	for pedimentopagado in pedimentospagados:
		print("{} - {} - {} - {} - {} - {} - {}".format(pedimentopagado.numeropedimento, 
			pedimentopagado.movimiento, 
			pedimentopagado.clavedocumento, 
			pedimentopagado.rfc, 
			pedimentopagado.archivom,
			pedimentopagado.banco.firmabancaria,
			pedimentopagado.banco.archivop
			))

def reportepedimentosvalidadospagados():
	"""Pedimentos validados y pagados"""
	pedimentosvalidadospagados = (Agencia_Aduanal
		.select(Agencia_Aduanal,Aduana,Banco)
		.join(Aduana, on=(Agencia_Aduanal.numeropedimento==Aduana.npedimento_validado))
		.switch(Agencia_Aduanal)
		.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
		)
	print("NumeroPedim - TP - CD - Re.Fed.Con. - Acuse Elc - Archivo Res - Firma Banc - FechaPago - Archivo Res")
	for pedimentovalidadopagado in pedimentosvalidadospagados:
		print("{} - {} - {} - {} - {} - {} - {} - {} - {}".format(pedimentovalidadopagado.numeropedimento, 
			pedimentovalidadopagado.movimiento, 
			pedimentovalidadopagado.clavedocumento, 
			pedimentovalidadopagado.rfc, 
			pedimentovalidadopagado.aduana.acuse,
			pedimentovalidadopagado.aduana.archivof,
			pedimentovalidadopagado.banco.firmabancaria,
			pedimentovalidadopagado.banco.fechapago,
			pedimentovalidadopagado.banco.archivop
			))

def exportarexcel():
	"""Exporta el consentrado de pedimentos a excel"""
	print("Espere un momento...")
	libro = xlsxwriter.Workbook("ReporteOperacionesS" + semana_anterior + ".xlsx")
	reporte = libro.add_worksheet("ReporteOperacion")
	fila = 0
	pedimentosvalidadospagados = (Agencia_Aduanal
		.select(Agencia_Aduanal,Aduana,Banco)
		.join(Aduana, on=(Agencia_Aduanal.numeropedimento==Aduana.npedimento_validado))
		.switch(Agencia_Aduanal)
		.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
		.where(Agencia_Aduanal.clavedocumento != "R1")
		)
	for pedimentovalidadopagado in pedimentosvalidadospagados:
		campos = [pedimentovalidadopagado.numeropedimento,
			pedimentovalidadopagado.movimiento,
			pedimentovalidadopagado.clavedocumento,
			pedimentovalidadopagado.rfc,
			pedimentovalidadopagado.aduana.acuse,
			pedimentovalidadopagado.aduana.archivof,
			pedimentovalidadopagado.banco.firmabancaria,
			pedimentovalidadopagado.banco.fechapago,
			pedimentovalidadopagado.banco.archivop
		]
		for i in range(len(campos)):
			reporte.write(fila , i, campos[i])
		fila = fila+1
	libro.close()
	time.sleep(.10)
	print("Exportacion finalizada....")

def exportarpagadoexcel():
	"""Exporta el consentrado de pedimentos pagados a excel"""
	print("Espere un momento...")
	libro = xlsxwriter.Workbook("ReportePagadosS" + semana_anterior + ".xlsx")
	reporte = libro.add_worksheet("ReportePagados")
	# Foramtos para las hojas de EXCEL
	formato_encabezado = libro.add_format({'bold':'true','align': 'center','valign': 'vcenter'})
	formato_encabezado_relleno = libro.add_format({'bold':'true','font_color':'white','align': 'center','valign': 'vcenter','fg_color': '#002060'})
	formato_subencabezado_relleno = libro.add_format({'bold':'true','font_color':'white','align': 'center','valign': 'vcenter','fg_color': '#1F4E78'})
	# Encabezado de las hojas de EXCEL
	reporte.insert_image('A3','logo.png')
	reporte.merge_range('D2:G5','REPORTE DE VALIDACIONES SEMANA '+ semana_anterior ,formato_encabezado)
	reporte.merge_range('F9:H9','AAALAC 4 048',formato_encabezado)
	reporte.merge_range('A11:H11','PERIODO '+fechareporteinicial+' A '+fechareportefinal,formato_encabezado_relleno)
	reporte.write('A12','PEDIMENTO', formato_subencabezado_relleno)
	reporte.write('B12','OPERACION', formato_subencabezado_relleno)
	reporte.write('C12','CLAVE DOC', formato_subencabezado_relleno)
	reporte.write('D12','RFC', formato_subencabezado_relleno)
	reporte.write('E12','ARCHIVO V', formato_subencabezado_relleno)
	reporte.write('F12','PATENTE', formato_subencabezado_relleno)
	reporte.write('G12','FECHA', formato_subencabezado_relleno)
	reporte.write('H12','ARCHIVO P', formato_subencabezado_relleno)
	# Ajustes de las columnas 
	reporte.set_column(0,0,13)
	reporte.set_column(1,2,10)
	reporte.set_column(3,3,16)
	reporte.set_column(4,5,11)
	reporte.set_column(6,6,9)
	reporte.set_column(7,7,11)
	# Autofiltros de los conceptos
	reporte.autofilter('A12:H12')
	pedimentospagados = (Agencia_Aduanal
					.select(Agencia_Aduanal, Banco)
					.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
					.where(Agencia_Aduanal.clavedocumento != "R1")
					)
	listacampoclave = []
	fila = 12 
	for pedimentopagado in pedimentospagados:
		if pedimentopagado.numeropedimento not in listacampoclave:
			campos = [pedimentopagado.numeropedimento, 
				pedimentopagado.movimiento, 
				pedimentopagado.clavedocumento, 
				pedimentopagado.rfc, 
				pedimentopagado.archivom,
				pedimentopagado.banco.patenteb,
				pedimentopagado.banco.fechapago,
				pedimentopagado.banco.archivop
			]
			listacampoclave.append(pedimentopagado.numeropedimento)
			for i in range(len(campos)):
				reporte.write(fila , i, campos[i])
			fila = fila+1
	libro.close()
	time.sleep(.10)
	print("Exportacion finalizada....")


def exportarestadisticasexcel():
	"""Exporta estadisticas a EXCEL semana"""
	print("Espere un momento...")
	libro = xlsxwriter.Workbook("ReporteEstadisticaS" + semana_anterior + ".xlsx")
	reporte = libro.add_worksheet("ReporteEstadistica")
	estadistica_hoja = libro.add_worksheet("Estadistica")
	grafico_barra_operacion = libro.add_chart({'type':'column'})
	fila = 0
	pedimentospagados = (Agencia_Aduanal
					.select(Agencia_Aduanal, Banco)
					.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
					.where(Agencia_Aduanal.clavedocumento != "R1")
					)
	listacampoclave = []
	patentesvalidaron = []
	estadistica = {}
	for pedimentopagado in pedimentospagados:
		if pedimentopagado.numeropedimento not in listacampoclave:
			campos = [pedimentopagado.numeropedimento, 
				pedimentopagado.movimiento, 
				pedimentopagado.clavedocumento, 
				pedimentopagado.banco.patenteb
			]
			listacampoclave.append(pedimentopagado.numeropedimento)
			patentesvalidaron.append(pedimentopagado.banco.patenteb)
			for i in range(len(campos)):
				reporte.write(fila , i, campos[i])
			fila = fila+1
	listapatente = set(patentesvalidaron)
	for x in listapatente:
		campos = [x,
		patentesvalidaron.count(x)]
		estadistica[str(x)] = {"Operacion":patentesvalidaron.count(x),"Importacion":0, "Exportacion":0, "TransitosNa":0, "TransitoIn":0}
	listacampoclave = []
	for pedimentopagado in pedimentospagados:
		if pedimentopagado.numeropedimento not in listacampoclave:
			listacampoclave.append(pedimentopagado.numeropedimento)
			for p in estadistica:
				if str(pedimentopagado.banco.patenteb) == str(p):
					if pedimentopagado.movimiento == "Impo":
						estadistica[p]["Importacion"] += 1
					elif pedimentopagado.movimiento == "Expo":
						estadistica[p]["Exportacion"] += 1
					if pedimentopagado.clavedocumento == "T3" and pedimentopagado.movimiento == "Impo":
						estadistica[p]["TransitosNa"] += 1
						estadistica[p]["Importacion"] -= 1
					elif pedimentopagado.clavedocumento == "T3" and pedimentopagado.movimiento == "Expo":
						estadistica[p]["TransitosNa"] += 1
						estadistica[p]["Exportacion"] -= 1						
					elif pedimentopagado.clavedocumento == "T7" and pedimentopagado.movimiento == "Impo":
						estadistica[p]["TransitoIn"] += 1
						estadistica[p]["Importacion"] -= 1
					elif pedimentopagado.clavedocumento == "T7" and pedimentopagado.movimiento == "Expo":
						estadistica[p]["TransitoIn"] += 1
						estadistica[p]["Exportacion"] -= 1
	fila = 1
	fila_b = 40
	columna = 10
	total_patente = len(listapatente)
	total_general = len(listacampoclave)
	topten = {}
	for (k,v) in sorted(estadistica.items()):
		cant_op = list(v.values())
		total_pate_oper = cant_op[0]
		topten[k] = total_pate_oper
	jk = total_patente
	for k,v in sorted(topten.items(), key=operator.itemgetter(1)):
		topten[k] = jk
		jk -= 1
	for (k,v) in sorted(estadistica.items()):
		grafico_barra = libro.add_chart({'type':'column'})
		grafico_pastel = libro.add_chart({'type':'pie'})
		cantidad = list(v.values())
		op = cantidad[0]
		im = cantidad[1]
		ex = cantidad[2]
		tn = cantidad[3]
		ti = cantidad[4]
		print("P:{} O:{:02d} I:{:02d} E:{:02d} TN:{:02d} TI:{:02d}".format(k, op, im, ex, tn, ti))
		campos_est = (k,op,im,ex,tn,ti,op,total_general)
		estadistica_hoja.write('A1',"PATENTE")
		estadistica_hoja.write('B1',"TOTAL_O")
		estadistica_hoja.write('C1',"IMPO")
		estadistica_hoja.write('D1',"EXPO")
		estadistica_hoja.write('E1',"TRAN_N")
		estadistica_hoja.write('F1',"TRAN_I")
		estadistica_hoja.write('G1',"CANTIDAD")
		estadistica_hoja.write('H1',"TOTAL")
		enca_pastel = str( topten[k] ) +"/"+ str( total_patente )
		porcentaje = (op / total_general)*100
		porcentaje = str ( round(porcentaje) )+"%"
		for i in range(len(campos_est)):
			estadistica_hoja.write(fila , i, campos_est[i])
		grafico_barra.add_series({'name':"Patente: "+ k,'categories':['Estadistica',0,2,0,5],'values':['Estadistica',fila,2,fila,5],})
		grafico_pastel.add_series({'name':"Patente: "+ k +" "+ porcentaje +" "+enca_pastel ,'categories':['Estadistica',0,6,0,7],'values':['Estadistica',fila,6,fila,7],})
		estadistica_hoja.insert_chart(fila,columna,grafico_barra)
		estadistica_hoja.insert_chart(fila_b,columna,grafico_pastel)
		fila = fila+1
		fila_b = fila_b+1
		columna = columna+1
	grafico_barra_operacion.add_series({'name':['Estadistica',0,1,0,1],'categories':['Estadistica',0,0,fila,0],'values':['Estadistica',1,1,fila,1],})
	estadistica_hoja.insert_chart('H1',grafico_barra_operacion)
	libro.close()
	pausa = input("Precione una tecla para continuar")
	time.sleep(.10)
	print("Exportacion finalizada....")


def exportarestadisticamensualsexcel():
	"""Exporta estadisticas del mes a excel"""
	print("Espere un momento...")
	libro = xlsxwriter.Workbook("ReporteEstadisticaS" + str(ano) + ".xlsx")
	reporte = libro.add_worksheet("ReporteEstadistica")
	estadistica_hoja = libro.add_worksheet("Estadistica")
	pedimentospagados = (Agencia_Aduanal
					.select(Agencia_Aduanal, Banco)
					.join(Banco, on=(Agencia_Aduanal.numeropedimento==Banco.npedimento_pagado))
					.where(Agencia_Aduanal.clavedocumento != "R1")
					)
	listacampoclave = []
	patentesvalidaron = []
	estadistica = {}
	mes = {1:'Enero', 2:'Febrero', 3:'Marzo', 4:'Abril', 5:'Mayo', 6:'Junio', 7:'Julio', 8:'Agosto', 9:'Septiembre', 10:'Octubre', 11:'Noviembre', 12:'Diciembre'}
	me = 0
	for m in range(len(mes)):
		fila = 0
		me += 1 
		print(mes[m])
		nombrehoja = "Estad_"+mes[m]+"_"+str(ano)
		hoja_mes = libro.add_sheet(nombrehoja)
		for pedimentopagado in pedimentospagados:
			mesoperacion = pedimentopagado.banco.fechapago[2:4]
			if int(mesoperacion) == me:
				if pedimentopagado.numeropedimento not in listacampoclave:
					campos = [pedimentopagado.numeropedimento, 
						pedimentopagado.movimiento, 
						pedimentopagado.clavedocumento, 
						pedimentopagado.banco.patenteb]
					listacampoclave.append(pedimentopagado.numeropedimento)
					patentesvalidaron.append(pedimentopagado.banco.patenteb)
				for i in range(len(campos)):
					hoja_mes.write(fila , i, campos[i])
				fila = fila+1
	listapatente = set(patentesvalidaron)
	for x in listapatente:
		campos = [x,
		patentesvalidaron.count(x)]
		estadistica[str(x)] = {"Operacion":patentesvalidaron.count(x),
		"Importacion":0, 
		"Exportacion":0, 
		"TransitosNa":0, 
		"TransitoIn":0}
	listacampoclave = []
	for pedimentopagado in pedimentospagados:
		if pedimentopagado.numeropedimento not in listacampoclave:
			listacampoclave.append(pedimentopagado.numeropedimento)
			for p in estadistica:
				if str(pedimentopagado.banco.patenteb) == str(p):
					if pedimentopagado.movimiento == "Impo":
						estadistica[p]["Importacion"] += 1
					elif pedimentopagado.movimiento == "Expo":
						estadistica[p]["Exportacion"] += 1
					if pedimentopagado.clavedocumento == "T3" and pedimentopagado.movimiento == "Impo":
						estadistica[p]["TransitosNa"] += 1
						estadistica[p]["Importacion"] -= 1
					elif pedimentopagado.clavedocumento == "T3" and pedimentopagado.movimiento == "Expo":
						estadistica[p]["TransitosNa"] += 1
						estadistica[p]["Exportacion"] -= 1						
					elif pedimentopagado.clavedocumento == "T7" and pedimentopagado.movimiento == "Impo":
						estadistica[p]["TransitoIn"] += 1
						estadistica[p]["Importacion"] -= 1
					elif pedimentopagado.clavedocumento == "T7" and pedimentopagado.movimiento == "Expo":
						estadistica[p]["TransitoIn"] += 1
						estadistica[p]["Exportacion"] -= 1
	fila = 0
	for (k,v) in sorted(estadistica.items()):
		cantidad = list(v.values())
		op = cantidad[0]
		im = cantidad[1]
		ex = cantidad[2]
		tn = cantidad[3]
		ti = cantidad[4]
		print("P:{} O:{:04d} I:{:04d} E:{:04d} TN:{:04d} TI:{:04d}".format(k, op, im, ex, tn, ti))
		campos_est = (k,op,im,ex,tn,ti)
		for i in range(len(campos_est)):
			estadistica_hoja.write(fila , i, campos_est[i])
		fila = fila+1
	pausa = input("Precione una tecla para continuar")
	time.sleep(.10)
	libro.close()
	print("Exportacion finalizada....")

menu = OrderedDict([
	('v',leer_variosj),
	('b',buscar_pedimento),
	('p',reportepatente),
	('a',reportepedimentovalidado),
	('l',reportepedimentopagado),
	('t',reportepedimentosvalidadospagados),
	('x',exportarexcel),
	('e',exportarpagadoexcel),
	('s',exportarestadisticasexcel),
	('m',exportarestadisticamensualsexcel)
])

if __name__ == '__main__':
	creacion_conexion()
	print()
	print("++++++++++++ Bienvenido al sistema de reportes V5 +++++++++++++++++")
	menu_loop()
	print()
