# -*- coding: utf-8 -*-
"""
Autor: Daniel Gómez
Revisó: Danilo Rodríguez
Aprobó: Juan Carlos Alvarez
versión 0.0

"""

import powerfactory as pf
import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell

##### Inicia la aplicación #####
app=pf.GetApplication()
app.ClearOutputWindow()
app.EchoOff()

##### Obtener los datos dentro del ComPython ######
script=app.GetCurrentScript()
Tipo_corto=script.Tipo_corto
Directorio=script.Directorio

sPhase=script.sPhase	  #### Fase fallada (Falla monofásica)
bPhase=script.bPhase      #### Fases falladas

filename=script.file_name

###### Fases falladas por defecto ############
if bPhase not in (0,1,2):
	app.PrintInfo('Seleccione una fase fallada válida, se ejecutará la falla sobre la fase A')
	bPhase=0
if sPhase not in (0,1,2):
	sPhase=0


#### Casos de estudio ###
shc_cases=script.GetContents('Cases')[0].GetAll('IntScenario') 

if shc_cases==[]:
	app.PrintError('No hay casos de estudio seleccionados para ejecutar el script, se ejecutará el caso activo')

### Crear el archivo de excel
try:
	writer=pd.ExcelWriter(Directorio+'\\'+filename+'.xlsx',engine='xlsxwriter') 
except:
	app.PrintError('Seleccione un directorio válido para guardar los resultados!!')


##### Objeto encargado de realizar los cortos circuitos
class Shc_cases:

	def __init__(self):
		###### Parámetros por defecto del corto circuito a ejecutar #####
		self.script=app.GetCurrentScript()
		self.Tipo_corto=self.script.Tipo_corto
		self.Directorio=self.script.Directorio
		self.shc_cases=self.script.GetContents('Cases')[0].GetAll('IntScenario')
		self.Barras=self.script.GetContents('Elementos')[0].GetAll('ElmTerm')
		self.shc=self.script.GetContents('Sch')[0]
		self.script.GetContents('Elementos')[0].Clear()
		self.script.GetContents('Elementos')[0].AddRef(self.Barras)
		self.fase_rara=0
		# self.Tipo_corto=3 		#### Método completo
		self.Tipo_falla='spgf'  #### Falla monofásica a tierra
		self.iopt_cur=0			#### Máxima corriente de corto circuito	


	##### Variables a devolver ####
	@property
	def I0x3(self):
		return self._I0x3

	@property
	def Ikss(self):
		return self._Ikss

	@property
	def IkssA(self):
		return self._IkssA

	@property
	def IkssB(self):
		return self._IkssB
	
	@property
	def IkssC(self):
		return self._IkssC

	@property
	def nombre(self):
		return self._nombre
	@property
	def tension(self):
		return self._tension
	
	
	###### Activar corto circuito ----- Extraer resultados ###
	def short_circuit(self,shc_case):
		shc_case.Activate()		### Activar el caso de estudio
		ldf=app.GetFromStudyCase('ComLdf')
		ldf.iopt_net=0
		self.shc.iopt_mde=self.Tipo_corto  #### Método de cortocircuito
		self.shc.iopt_cur=self.iopt_cur
		self.shc.iopt_allbus=0
		# if self.Tipo_corto==3:
		# 	self.shc.c_ldf.iopt_net=0 ### Forzar flujo de carga balanceado

		self.shc.iopt_shc=self.Tipo_falla
		self.shc.shcobj=self.script.GetContents('Elementos')[0]

		if self.Tipo_falla=='2psc':
			self.shc.i_p2psc=self.fase_rara

		elif self.Tipo_falla=='2pgf':
			self.shc.i_p2pgf=self.fase_rara	

		elif self.Tipo_falla=='spgf':
			self.shc.i_pspgf=self.fase_rara



		self.shc.Execute()       ### Correr el corto circuito

		#### Extrae el nombre de las barras falladas ####
		self._nombre=[]
		for self.Barra in self.Barras:
			# self._nombre.append(self.Barra.loc_name)
			try:
				self._nombre.append(self.Barra.cpSubstat.loc_name+" "+self.Barra.loc_name)
			except:
				self._nombre.append(self.Barra.loc_name)


		#### Extrae la tensión
		self._tension=[self.Barra.uknom for self.Barra in self.Barras] 


		#### Extrae corriente 3I0 y Ikss para fallas monofásicas	
		if self.Tipo_falla=='spgf':

			self._I0x3=[]
			self._Ikss=[]
			for self.Barra in self.Barras:
				try:
					self._I0x3.append(self.Barra.GetAttribute('m:I0x3'))
				except:
					self._I0x3.append(0)

				try:
					self._Ikss.append(self.Barra.GetAttribute('m:Ikss'))
				except:
					self._Ikss.append(0)


		##### Ikss para fallas trifásicas
		if self.Tipo_falla=='3psc':
			self._Ikss=[]
			for self.Barra in self.Barras:
				try:
					self._Ikss.append(self.Barra.GetAttribute('m:Ikss'))
					# self._Ikss=[self.Barra.GetAttribute('m:Ikss') for self.Barra in self.Barras ]
				except:
					self._Ikss.append(0)

		##### Extraer datos fallas bifásicas aisladas
		if self.Tipo_falla=='2psc':
			self._IkssA=[]
			self._IkssB=[]
			self._IkssC=[]
			for self.Barra in self.Barras:
				try:
					self._IkssA.append(self.Barra.GetAttribute('m:Ikss:A'))
				except:
					self._IkssA.append(0)
				try:
					self._IkssB.append(self.Barra.GetAttribute('m:Ikss:B'))
				except:
					self._IkssB.append(0)
				try:
					self._IkssC.append(self.Barra.GetAttribute('m:Ikss:C'))
				except:
					self._IkssC.append(0)

		#### Extraer datos de fallas bifásicas a tierra ####
		if self.Tipo_falla=='2pgf':
			self._IkssA=[]
			self._IkssB=[]
			self._IkssC=[]
			self._I0x3=[]
			for self.Barra in self.Barras:
				try:
					self._IkssA.append(self.Barra.GetAttribute('m:Ikss:A'))
				except:
					self._IkssA.append(0)
				try:
					self._IkssB.append(self.Barra.GetAttribute('m:Ikss:B'))
				except:
					self._IkssB.append(0)
				try:
					self._IkssC.append(self.Barra.GetAttribute('m:Ikss:C'))
				except:
					self._IkssC.append(0)
				try:
					self._I0x3.append(self.Barra.GetAttribute('m:I0x3'))
				except:
					self._I0x3.append(0)


### Guardar datos corto monofásico en un dataframe
def corto_mono_keep(Nombre,I0x3='',Ikss=''):
	datos_mono=pd.DataFrame()
	datos_mono['Nombre']=list(Nombre)
	# try:
	# 	datos_mono['Ikss']=list(Ikss)
	# except:
	# 	pass
	try:
		datos_mono['3I0']=list(I0x3)
	except:
		pass
	datos_mono=datos_mono.sort_values(by='Nombre')
	datos_mono=datos_mono.drop_duplicates(subset='Nombre')
	datos_mono=datos_mono.drop(['Nombre'],axis=1)
	return(datos_mono)

### Guardar datos corto bifásico a tierra en un dataframe
def corto_bg_keep(Nombre,I0x3,IkssA,IkssB,IkssC):
	datos_corto_bg=pd.DataFrame()
	datos_corto_bg['Nombre']=list(Nombre)
	# try:
	# 	datos_corto_bg['Ikss']=list(IkssA)
	# 	# app.PrintPlain('El promedio de la fila A es:'+str(datos_corto_bg['Ikss A'].mean()))
	# 	if datos_corto_bg['Ikss'].mean()==0:
	# 		datos_corto_bg.drop(['Ikss A'],axis=1)
	# 		datos_corto_bg=datos_corto_bg.drop(['Ikss A'],axis=1)
	# except:
	# 	pass
	# try:
	# 	datos_corto_bg['Ikss B']=list(IkssB)
	# 	# app.PrintPlain('El promedio de la fila B es:'+str(datos_corto_bg['Ikss B'].mean()))
	# 	if datos_corto_bg['Ikss B'].mean()==0:
	# 		datos_corto_bg.drop(['Ikss B'],axis=1)
	# 		datos_corto_bg=datos_corto_bg.drop(['Ikss B'],axis=1)
	# except:
	# 	pass
	# try:
	# 	datos_corto_bg['Ikss C']=list(IkssC)
	# 	# app.PrintPlain('El promedio de la fila C es:'+str(datos_corto_bg['Ikss C'].mean()))
	# 	if datos_corto_bg['Ikss C'].mean()==0:
	# 		datos_corto_bg.drop(['Ikss C'],axis=1)
	# 		datos_corto_bg=datos_corto_bg.drop(['Ikss C'],axis=1)
	# except:
	# 	pass
	
	datos_corto_bg['3I0']=list(I0x3)
	datos_corto_bg=datos_corto_bg.sort_values(by='Nombre')
	datos_corto_bg=datos_corto_bg.drop_duplicates(subset='Nombre')
	datos_corto_bg=datos_corto_bg.drop(['Nombre'],axis=1)
	return(datos_corto_bg)

### Guardar datos corto bifásico en un dataframe
def corto_bi_keep(Nombre,IkssA='',IkssB='',IkssC=''):
	datos_corto_bi=pd.DataFrame()
	datos_corto_bi['Nombre']=list(Nombre)
	try:
		datos_corto_bi['Ikss A']=list(IkssA)
		if datos_corto_bi['Ikss A'].mean()==0:
			datos_corto_bi.drop(['Ikss A'],axis=1)
			datos_corto_bi=datos_corto_bi.drop(['Ikss A'],axis=1)
	except:
		pass
	try:
		datos_corto_bi['Ikss B']=list(IkssB)
		if datos_corto_bi['Ikss B'].mean()==0:
			datos_corto_bi.drop(['Ikss B'],axis=1)
			datos_corto_bi=datos_corto_bi.drop(['Ikss B'],axis=1)
	except:
		pass
	try:
		datos_corto_bi['Ikss C']=list(IkssC)
		if datos_corto_bi['Ikss C'].mean()==0:
			datos_corto_bi.drop(['Ikss C'],axis=1)
			datos_corto_bi=datos_corto_bi.drop(['Ikss C'],axis=1)
	except:
		pass

	datos_corto_bi=datos_corto_bi.sort_values(by='Nombre')
	datos_corto_bi=datos_corto_bi.drop_duplicates(subset='Nombre')
	datos_corto_bi=datos_corto_bi.drop(['Nombre'],axis=1)
	return(datos_corto_bi[datos_corto_bi.columns[0]])

### Guardar datos corto trifásico en un dataframe
def corto_tri_keep(Nombre,Ikss,tension):
	datos_corto_tri=pd.DataFrame()
	datos_corto_tri['Nombre']=list(Nombre)
	datos_corto_tri['Tensión [kV]']=list(tension)
	try:
		datos_corto_tri['Ikss [kA]']=list(Ikss)
	except:
		datos_corto_tri['Ikss [kA]']=''
	datos_corto_tri=datos_corto_tri.sort_values(by='Nombre')
	datos_corto_tri=datos_corto_tri.drop_duplicates(subset='Nombre')

	return(datos_corto_tri)

### unir dataframes
def join_dataframe(data1,dato2,dato3,dato4):
	joined_dataframe=pd.concat([data1,dato2,dato3,dato4], axis=1)
	return(joined_dataframe)
	
### Exportar datos a excel con el respectivo formato
def export_data(case,dataframe,startrow=2,header=''):
	try:
		pd.formats.format.header_style = None
	except:
		pass
	dataframe.to_excel(writer, sheet_name=str(case), startrow=startrow ,index=False)
	workbook=writer.book
	worksheet=writer.sheets[str(case)]
	worksheet.set_zoom(80)
	header_format=workbook.add_format({'align':'center', 'bold':True})
	header_format.set_font_name('Arial')
	parcial_fmt=workbook.add_format({'align':'center', 'num_format':'0.0000'})
	parcial_fmt.set_font_name('Arial')
	parcial_fmt2=workbook.add_format({'align':'center'})
	parcial_fmt2.set_font_name('Arial')

	parcial_fmt3=workbook.add_format({'align':'center','num_format':'0.00'})
	parcial_fmt3.set_font_name('Arial')

	worksheet.set_column('A:A',27,parcial_fmt2)
	worksheet.set_column('B:B',16,parcial_fmt3)
	worksheet.set_column('C:K',16,parcial_fmt)
	worksheet.set_row(startrow,None,header_format)

	merge_format = workbook.add_format({
    'bold': 1,
    'align': 'center',
    'valign': 'vcenter'})
	merge_format.set_font_name('Arial')
	merge_format.set_text_wrap()

	worksheet.merge_range('A1:F1', 'Corrientes máximas', merge_format)
	# worksheet.merge_range('A{}:F{}'.format((startrow-1),(startrow-1)), 'Corrientes mínimas', merge_format)

	worksheet.write('D2:D2', 'Falla 2\u03C6 [kA]', merge_format)
	worksheet.write('E2:E2', 'Falla 2\u03C6-T [kA]', merge_format)
	worksheet.write('F2:F2', 'Falla 1\u03C6 [kA]', merge_format)


	worksheet.merge_range('D{}:D{}'.format((startrow),(startrow)), 'Falla 2\u03C6 [kA]', merge_format)
	worksheet.merge_range('E{}:E{}'.format((startrow),(startrow)), 'Falla 2\u03C6-T [kA]', merge_format)
	worksheet.merge_range('F{}:F{}'.format((startrow),(startrow)), 'Falla 1\u03C6 [kA]', merge_format)

	worksheet.merge_range('A2:A3', 'Nombre', merge_format)
	worksheet.merge_range('B2:B3', 'Tensión [kV]', merge_format)
	worksheet.merge_range('C2:C3', 'Falla 3\u03C6 Ikss [kA]', merge_format)

	worksheet.merge_range('A{}:A{}'.format((startrow),(startrow+1)), 'Nombre', merge_format)
	worksheet.merge_range('B{}:B{}'.format((startrow),(startrow+1)), 'Tensión [kV]', merge_format)

	worksheet.merge_range('C{}:C{}'.format((startrow),(startrow+1)), 'Falla 3\u03C6 Ikss [kA]', merge_format)



################### Cortos Máximos #######################
Corto_mono_max=Shc_cases()         ### monofásico
Corto_mono_max.fase_rara=sPhase

Corto_bg_max=Shc_cases()			#### Bifásico a tierra
Corto_bg_max.Tipo_falla='2pgf'
Corto_bg_max.fase_rara=bPhase

Corto_bi_max=Shc_cases()			#### Bifásico aislado
Corto_bi_max.Tipo_falla='2psc'
Corto_bi_max.fase_rara=bPhase

Corto_tri_max=Shc_cases()			#### Trifásico
Corto_tri_max.Tipo_falla='3psc'


################### Cortos Mínimos #######################
Corto_mono_min=Shc_cases()		### monofásico
Corto_mono_min.iopt_cur=1
Corto_mono_min.fase_rara=sPhase

Corto_bg_min=Shc_cases()		#### Bifásico a tierra
Corto_bg_min.Tipo_falla='2pgf'
Corto_bg_min.iopt_cur=1
Corto_bg_min.fase_rara=bPhase

Corto_bi_min=Shc_cases()		#### Bifásico aislado
Corto_bi_min.Tipo_falla='2psc'
Corto_bi_min.iopt_cur=1
Corto_bi_min.fase_rara=bPhase

Corto_tri_min=Shc_cases()		#### Trifásico
Corto_tri_min.Tipo_falla='3psc'
Corto_tri_min.iopt_cur=1



######### Ejecución sobre los casos de estudio ###########

##### En caso de que no se seleccione ningún caso de estudio, se ejecuta sobre el que esté activo#####

if shc_cases==[]:

	case=app.GetActiveStudyCase()

	Corto_mono_max.short_circuit(case)			### Función para guardar datos
	datos_Corto_mono_max=corto_mono_keep(Corto_mono_max.nombre,Corto_mono_max.I0x3,Corto_mono_max.Ikss)

	Corto_bg_max.short_circuit(case)
	datos_Corto_bg_max=corto_bg_keep(Corto_bg_max.nombre,Corto_bg_max.I0x3,Corto_bg_max.IkssA,Corto_bg_max.IkssB,Corto_bg_max.IkssC)

	Corto_bi_max.short_circuit(case)
	datos_Corto_bi_max=corto_bi_keep(Corto_bi_max.nombre,Corto_bi_max.IkssA,Corto_bi_max.IkssB,Corto_bi_max.IkssC)

	Corto_tri_max.short_circuit(case)
	datos_Corto_tri_max=corto_tri_keep(Corto_tri_max.nombre,Corto_tri_max.Ikss,Corto_tri_max.tension)
		###### Unir datos
	datos_maximos=join_dataframe(datos_Corto_tri_max,datos_Corto_bi_max,datos_Corto_bg_max,datos_Corto_mono_max)
	if len(case.loc_name)>31:
		export_data(case.loc_name[:30],datos_maximos)
	else:
		export_data(case.loc_name,datos_maximos)


	# ######### Cortos Mínimos ########

	# Corto_mono_min.short_circuit(case)
	# datos_Corto_mono_min=corto_mono_keep(Corto_mono_min.nombre,Corto_mono_min.I0x3,Corto_mono_min.Ikss)

	# Corto_bg_min.short_circuit(case)
	# datos_Corto_bg_min=corto_bg_keep(Corto_bg_min.nombre,Corto_bg_min.I0x3,Corto_bg_min.IkssA,Corto_bg_min.IkssB,Corto_bg_min.IkssC)

	# Corto_bi_min.short_circuit(case)
	# datos_Corto_bi_min=corto_bi_keep(Corto_bi_min.nombre,Corto_bi_min.IkssA,Corto_bi_min.IkssB,Corto_bi_min.IkssC)

	# Corto_tri_min.short_circuit(case)
	# datos_Corto_tri_min=corto_tri_keep(Corto_tri_min.nombre,Corto_tri_min.Ikss,Corto_tri_min.tension)

	# datos_minimos=join_dataframe(datos_Corto_tri_min,datos_Corto_bi_min,datos_Corto_bg_min,datos_Corto_mono_min)
	# if len(case.loc_name)>30:
	# 	export_data(case.loc_name[:30],datos_minimos, startrow=(datos_maximos.shape[0]+7))
	# else:
	# 	export_data(case.loc_name,datos_minimos, startrow=(datos_maximos.shape[0]+7))

else:
	for case in shc_cases:
		########## Cortos máximos ########
		app.PrintPlain(case)
		Corto_mono_max.short_circuit(case)			### Función para guardar datos
		datos_Corto_mono_max=corto_mono_keep(Corto_mono_max.nombre,Corto_mono_max.I0x3,Corto_mono_max.Ikss)

		Corto_bg_max.short_circuit(case)
		datos_Corto_bg_max=corto_bg_keep(Corto_bg_max.nombre,Corto_bg_max.I0x3,Corto_bg_max.IkssA,Corto_bg_max.IkssB,Corto_bg_max.IkssC)

		Corto_bi_max.short_circuit(case)
		datos_Corto_bi_max=corto_bi_keep(Corto_bi_max.nombre,Corto_bi_max.IkssA,Corto_bi_max.IkssB,Corto_bi_max.IkssC)

		Corto_tri_max.short_circuit(case)
		datos_Corto_tri_max=corto_tri_keep(Corto_tri_max.nombre,Corto_tri_max.Ikss,Corto_tri_max.tension)
			###### Unir datos
		datos_maximos=join_dataframe(datos_Corto_tri_max,datos_Corto_bi_max,datos_Corto_bg_max,datos_Corto_mono_max)
		if len(case.loc_name)>31:
			export_data(case.loc_name[:30],datos_maximos)
		else:
			export_data(case.loc_name,datos_maximos)


		######### Cortos Mínimos ########

		# Corto_mono_min.short_circuit(case)
		# datos_Corto_mono_min=corto_mono_keep(Corto_mono_min.nombre,Corto_mono_min.I0x3,Corto_mono_min.Ikss)

		# Corto_bg_min.short_circuit(case)
		# datos_Corto_bg_min=corto_bg_keep(Corto_bg_min.nombre,Corto_bg_min.I0x3,Corto_bg_min.IkssA,Corto_bg_min.IkssB,Corto_bg_min.IkssC)

		# Corto_bi_min.short_circuit(case)
		# datos_Corto_bi_min=corto_bi_keep(Corto_bi_min.nombre,Corto_bi_min.IkssA,Corto_bi_min.IkssB,Corto_bi_min.IkssC)

		# Corto_tri_min.short_circuit(case)
		# datos_Corto_tri_min=corto_tri_keep(Corto_tri_min.nombre,Corto_tri_min.Ikss,Corto_tri_min.tension)

		# datos_minimos=join_dataframe(datos_Corto_tri_min,datos_Corto_bi_min,datos_Corto_bg_min,datos_Corto_mono_min)
		# if len(case.loc_name)>30:
		# 	export_data(case.loc_name[:30],datos_minimos, startrow=(datos_maximos.shape[0]+7))
		# else:
		# 	export_data(case.loc_name,datos_minimos, startrow=(datos_maximos.shape[0]+7))


writer.save()
app.PrintInfo('Ejecución exitosa! diríjase a {}\corto_barras.xlsx para ver el archivo resultante'.format(Directorio)) 
