#! ../venv/bin/python
import sys
import pandas as pd
import xlutils
from keabook import Keabook

def main(argv):
	'''The main function.'''
	fileName			= argv[1]
	excelFile			= pd.ExcelFile(fileName)
	table				= excelFile.parse('Data')
	table				= table[['PlantID','Sample ID','Plant Alt Names','Plate No','Position on Plate(s)','Tray Number','Row','Column']]
	table				= table.dropna(how='all',subset=['Sample ID'])
	table				= addPopulationColumn(table)
	table				= table.dropna(how='all',subset=['Plate No'])
	table['Plate']		= table['Tray Number'].apply(getPlate)
	table['Tray']		= table['Plate'].apply(getTray)
	plates				= table[['Plate','Plate No','Tray']].drop_duplicates().reset_index(drop=True)
	
	keabook				= Keabook()
	keabook.addPlatesTable(plates)
	keabook.addHarvestSheets(plates)
	keabook.addHarvestWells(table)
	keabook.save('output.xls')

def addPopulationColumn(table):
	for r,row in table.iterrows():
		if row['PlantID'] == 'Population':
			population	= row['Sample ID']
		else:
			table.loc[r,'Population'] = population
	return table
	
def getPlate(trayNumber):
	return trayNumber.split(' ',1)[1][1:-1]
	
def getTray(plate):
	return plate.split(' ',1)[1]
	




if __name__=='__main__':
	main(sys.argv)
