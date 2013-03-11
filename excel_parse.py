import xml.sax
import sys
import math
import excel_lang
import pickle

class ExcelWB():

	def __init__(self):
		self.xlData = {}
		self.xlData["namedRanges"] = []
		self.xlData["namedCells"]  = {}
		self.xlData["worksheets"]  = {}

	def setCellData(self, worksheet, col, row, field, value):
		sCol = self.col2Name(col)
		sRow = str(row)
		if not worksheet in self.xlData["worksheets"]:
			self.xlData["worksheets"][worksheet] = {}
		if not sCol in self.xlData["worksheets"][worksheet]:
			self.xlData["worksheets"][worksheet][sCol] = {}
		if not sRow in self.xlData["worksheets"][worksheet][sCol]:
			self.xlData["worksheets"][worksheet][sCol][sRow] = {}
		if not field in self.xlData["worksheets"][worksheet][sCol][sRow]:
			self.xlData["worksheets"][worksheet][sCol][sRow][field] = value

	def getCell(self, worksheet, col, row, type):
		if not isinstance(col, str): col = col2Name(col)
		if not isinstance(row, str): row = str(row)
		return self.xlData["worksheets"][worksheet][col][str(row)][type]

	def getNamedCell(self, workbook, name):
		if name in self.xlData["namedCells"]:
			cellRange = self.xlData["namedCells"][name]
			if len(cellRange) == 1:
				print("named cell points to:", cellRange[0][0], cellRange[0][1], cellRange[0][2])
				cell = ExcelCell(cellRange[0][0], 
					cellRange[0][1],
					self.getCell(cellRange[0][0], cellRange[0][1], cellRange[0][2], 'content'),
					self.getCell(cellRange[0][0], cellRange[0][1], cellRange[0][2], 'formula')
				)
				return cell

	def setNamedCell(self, worksheet, cellName, col, row):
		if not cellName in self.xlData["namedCells"]:
			self.addNamedCell(cellName)
		self.xlData["namedCells"][cellName].append((worksheet, self.col2Name(col), str(row)))

	def addNamedRange(self, rangeName, rangeRefersTo):
		self.xlData["namedRanges"].append({'name':rangeName, 'refersto':rangeRefersTo})

	def addNamedCell(self, cellName):
		self.xlData["namedCells"][cellName] = []

	def col2Name(self, columnValue):
		'''Converts an integer column value to an Excel Alpha based name.'''
		columnValue -= 1  # 26 should yield Z, which is A + 25
		x = math.floor(columnValue / 26) - 1
		y = columnValue % 26
		BigA = ord('A')
		col = ""
		if x > 0:
			col = chr(BigA + x)
		return col + chr(BigA + y)

	def getData(self):
		return self.xlData

	def dump(self):
		print("WORKSHEETS\n", self.xlData["worksheets"])
		print("NAMED CELLS\n", self.xlData["namedCells"])
		print("NAMED RANGES\n", self.xlData["namedRanges"])


class ExcelCell():

	def __init__(self, col, row, data, formula=None):
		self.col = col
		self.row = row
		self.data = data[0]
		self.datatype = data[1]
		self.formula = formula

	def getAddress(self):
		return self.col + self.row

	def getFormula(self):
		return self.formula

	def getData(self):
		return [self.data, self.datatype]

	def getDataType(self):
		return self.datatype

	def calOffset(self, relcell):
		print(relcell)
		return



class ExcelHandler(xml.sax.ContentHandler):

	def __init__(self):
		xml.sax.ContentHandler.__init__(self)
		self.ok = False
		self.state = []

		self.wb = ExcelWB()

		self.row = 0
		self.col = 0

	def startElement(self, name, attrs):
		# print("startElement '" + name + "'")
		self.state.append(name)
		if self.ok:
			if name == 'Worksheet':
				self.worksheet = attrs['ss:Name']
				print("Worksheet: " + self.worksheet)
			elif name == 'NamedRange':
				self.wb.addNamedRange(attrs['ss:Name'], attrs['ss:RefersTo'])
				# self.xlData["namedRanges"].append({'name':attrs['ss:Name'], 'refersto':attrs['ss:RefersTo']})

			# a table marks the start of the worksheet's data  (rows and cells)
			elif name == 'Table':
				self.row = 0

			# a row is a collection of cells
			elif name == 'Row' and 'Table' in self.state:
				self.col = 0
				if 'ss:Index' in attrs:
					self.row = int(attrs['ss:Index'])
				else:
					self.row += 1

			# a cell is a holder of data and other goodies
			elif name == 'Cell' and 'Row' in self.state:
				if 'ss:Index' in attrs:
					self.col = int(attrs['ss:Index'])
				else:
					self.col += 1
				if 'ss:Formula' in attrs:
					self.wb.setCellData(self.worksheet, self.col, self.row, 'formula', attrs['ss:Formula'])

			# finally if we have a data element, then look at it
			elif name == 'Data' and 'Cell' in self.state:
				# print("got data")
				pass

			# save any named cell references
			elif name == 'NamedCell' and self.state[-2] == 'Cell':
				if 'ss:Name' in attrs:
					self.wb.setNamedCell(self.worksheet, attrs['ss:Name'], self.col, self.row)
					# self.xlData["namedCells"][attrs['ss:Name']].append((self.col2Name(self.col), str(self.row)))
				else:
					print("ERROR: A named cell element found without an ss:Name attribute...what to do?")
					sys.exit(100)

		elif name == 'Workbook':
			self.ok = True

	def endElement(self, name):
		if self.ok:
			stateName = self.state.pop()
			if name != stateName:
				print("ERROR: popping '" + stateName + "' but closing element '" + name + "'")
			if name == 'Worksheet':
				print('Worksheet closing: ' + self.worksheet)


	def characters(self, content):
		if 'Data' == self.state[-1]:
			self.wb.setCellData(self.worksheet, self.col, self.row, 'content', content)

	def terminate(self):
		self.wb.dump()
		sys.exit()

	def getWorkbook(self):
		return self.wb


def fetchCell(cell):
	example_string = cell.getFormula()
	print(example_string)

	if example_string:
		# Tokenize
		# excel_lang.lexer.input(example_string)
		# while True:
		#     tok = excel_lang.lexer.token()
		#     if not tok: break      # No more input
		#     print(tok)

		t = excel_lang.yacc.parse(example_string)
		expand(cell)


def expand(cell):

		if t[0] == 'FORMULA':
			print(t)
			args = t[2]
			print(t[1], args)

	else:
		print("No formula returned from that cell...")


def main(sourceFileName):

	try:
		xlpickle = open("xlstruct.pkl", "r+b")
		U = pickle.Unpickler(xlpickle)
		xlWB = U.load()

	except:
		xlpickle = open("xlstruct.pkl", "w+b")
		source = open(sourceFileName, 'r', encoding='utf-8')
		handler = ExcelHandler()
		xml.sax.parse(source, handler)

		xlWB = handler.getWorkbook()
		P = pickle.Pickler(xlpickle)
		P.dump(xlWB)

	# example_string = xlWB.getCell('PegTop', 'ES', '252', 'formula')
	cell = xlWB.getNamedCell(None, 'YResult')
	# example_string = "=IF(R[-277]C>R[-305]C-2*R[-303]C,'-',0.58*R[-319]C*PI()*R[-277]C*R[-303]C*R[-3]C/R[-404]C/1000)"
	
	fetchCell(cell)


if __name__ == '__main__':
	main('design_check.xml')
