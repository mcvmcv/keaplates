from xlwt import Workbook, easyxf

########################################################################
####	The Keabook class 											####
########################################################################
class Keabook(Workbook):
########################################################################
	def __init__(self):
		super(Keabook,self).__init__()
		self.sheetList						= []
		self.lookup3						= {'A':7,'B':10,'C':13,'D':16,'E':19,'F':22,'G':25,'H':28}
		self.lookup							= {'A':7,'B':11,'C':15,'D':19,'E':23,'F':27,'G':31,'H':35}
		self.black							= easyxf('pattern: pattern solid, fore_colour black; font: name Arial, bold True, colour white;')
		self.table							= easyxf('borders: left thin, right thin, top thin, bottom thin;')
		self.btop							= easyxf('pattern: pattern solid, fore_colour pale_blue; font: name Arial, bold True, colour black; borders: left thick, right thick, top thick;')
		self.bmid							= easyxf('pattern: pattern solid, fore_colour pale_blue; font: name Arial, bold True, colour black; borders: left thick, right thick,;')
		self.bbot							= easyxf('pattern: pattern solid, fore_colour pale_blue; font: name Arial, bold True, colour black; borders: left thick, right thick, bottom thick;')
		self.ctop							= easyxf('borders: left thick, right thick, top thick;')
		self.cmid							= easyxf('borders: left thick, right thick;')
		self.cbot							= easyxf('borders: left thick, right thick, bottom thick;')
		
		
########################################################################
	def addSheet(self,name):
		self.sheetList.append(name)
		return self.add_sheet(name)
			
	def getSheet(self,name):
		index					= self.sheetList.index(name)
		return self.get_sheet(index)
	
	def getOrCreateSheet(self,name):
		if name in self.sheetList:
			return self.getSheet(name)
		else:
			return self.addSheet(name)
			
########################################################################
	def addHarvestHeaders(self,sheet,row):
		'''Adds the info at the top of the plate required for harvest
		plate sheets.'''
		sheet.write(0,						0,	'Kea Plate',		self.black)
		sheet.write(0,						1,	row['Plate No'])
		sheet.write(1,						0,	'Tray',				self.black)
		sheet.write(1,						1,	row['Tray'])
		sheet.write(2,						0,	'Date',				self.black)
		sheet.write(2,						1,	'')
		sheet.write(3,						0,	'Harvester',		self.black)
		sheet.write(3,						1,	'')
		
	def addBorder(self,sheet):
		'''Adds the coloured border around the plate representation. The
		'colour' argument is a string naming an Excel colour.'''
		colLabels							= range(1,13)
		rowLabels							= 'ABCDEFGH'
		
		sheet.write(5,						0,	'',					self.btop)
		sheet.write(6,						0,	'',					self.bbot)
		for c in colLabels:
			sheet.write(5,					c,	'',					self.btop)
			sheet.write(6,					c,	c,					self.bbot)
		for r in rowLabels:
			sheet.write(self.lookup[r],		0,	r,					self.btop)
			sheet.write(self.lookup[r]+1,	0,	'',					self.bmid)
			sheet.write(self.lookup[r]+2,	0,	'',					self.bmid)
			sheet.write(self.lookup[r]+3,	0,	'',					self.bbot)
			
	def writeWell(self,sheet,well,data):
		'''Writes the information in data (a three-tuple) to well, with
		colour as the background colour.'''
		r									= self.lookup[well[0]]
		c									= int(well[1:])
		sheet.write(r,						c,	data[0],		self.ctop)
		sheet.write(r+1,					c,	data[1],		self.cmid)
		sheet.write(r+2,					c,	data[2],		self.cmid)
		sheet.write(r+3,					c,	data[3],		self.cbot)
			
########################################################################
	def addPlatesTable(self,plates):
		sheet										= self.addSheet('Plates')
		trayCol										= sheet.col(1)
		trayCol.width								= 256*20
		notesCol									= sheet.col(4)
		notesCol.width								= 256*30
		sheet.write(0,	0,	'Kea Plate ID',			self.black)
		sheet.write(0,	1,	'Tray Name',			self.black)
		sheet.write(0,	2,	'Harvester',			self.black)
		sheet.write(0,	3,	'Date',					self.black)
		sheet.write(0,	4,	'Notes',				self.black)
		
		for r,row in plates.iterrows():
			sheet.write(r+1,0,row['Plate No'],		self.table)
			sheet.write(r+1,1,row['Tray'],			self.table)
			sheet.write(r+1,2,'',					self.table)
			sheet.write(r+1,3,'',					self.table)
			sheet.write(r+1,4,'',					self.table)
	
########################################################################
	def addHarvestSheets(self,plates):
		for r, row in plates.iterrows():
			name				= row['Plate']
			sheet				= self.getOrCreateSheet(name)
			self.addHarvestHeaders(sheet,row)
			self.addBorder(sheet)
			for c in range(1,13):
				col				= sheet.col(c)
				col.width		= 256*11
			
########################################################################
	def addHarvestWells(self,samples):
		for r, row in samples.iterrows():
			plateName			= row['Plate']
			sheet				= self.getSheet(plateName)
			well				= row['Position on Plate(s)']
			data1				= row['Sample ID']
			data2				= row['Plant Alt Names'][4:]
			data3				= ''
			data4				= ''
			data				= (data1,data2,data3,data4)
			self.writeWell(sheet,well,data)




