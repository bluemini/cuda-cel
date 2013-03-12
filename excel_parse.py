import xml.sax
import sys
import math
import excel_lang
import pickle


'''
A cell is referenced as an instance of an ExcelCell
An address, to either fetch or as a property of a cell is a 3-tuple:
	(sheet, row, column)
'''

class ExcelWB():

	def __init__(self):
		self.xlData = {}
		self.xlData["namedRanges"] = []
		self.xlData["namedCells"]  = {}
		self.xlData["worksheets"]  = {}

	def setCellData(self, cell, field, value):
		'''each cell location in the nested dict of objects. The 'data' for the
		cell is a dict, with each key representing an associated piece of data, its
		name, type and other meta-data.'''
		cellAddress = cell.getAddress()
		col = cellAddress[2]
		row = str(cellAddress[1])
		worksheet = cellAddress[0]
		if not worksheet in self.xlData["worksheets"]:
			self.xlData["worksheets"][worksheet] = {}
		if not col in self.xlData["worksheets"][worksheet]:
			self.xlData["worksheets"][worksheet][col] = {}
		if not row in self.xlData["worksheets"][worksheet][col]:
			self.xlData["worksheets"][worksheet][col][row] = {}
		if not field in self.xlData["worksheets"][worksheet][col][row]:
			self.xlData["worksheets"][worksheet][col][row][field] = value

	def getCell(self, address): #, worksheet, col, row, type
		col = address[2]
		row = address[1]
		worksheet = address[0]
		cell = ExcelCell(worksheet, 
			col, 
			row, 
			self.xlData["worksheets"][worksheet][col][row]
		)
		return cell

	def getNamedCell(self, workbook, name):
		if name in self.xlData["namedCells"]:
			cellRange = self.xlData["namedCells"][name]
			if len(cellRange) == 1:
				print("named cell points to:", cellRange[0][0], cellRange[0][1], cellRange[0][2])
				return self.getCell( (cellRange[0][0], cellRange[0][2], cellRange[0][1]) )

	def setNamedCell(self, worksheet, cellName, col, row):
		if not cellName in self.xlData["namedCells"]:
			self.addNamedCell(cellName)
		self.xlData["namedCells"][cellName].append((worksheet, self.col2Name(col), str(row)))

	def addNamedRange(self, rangeName, rangeRefersTo):
		self.xlData["namedRanges"].append({'name':rangeName, 'refersto':rangeRefersTo})

	def addNamedCell(self, cellName):
		self.xlData["namedCells"][cellName] = []

	def getData(self):
		return self.xlData

	def dump(self):
		print("WORKSHEETS\n", self.xlData["worksheets"])
		print("NAMED CELLS\n", self.xlData["namedCells"])
		print("NAMED RANGES\n", self.xlData["namedRanges"])


class ExcelCell():

	def __init__(self, sheet, col, row, data):
		self.sheet = sheet
		self.col = col
		self.row = row
		self.data = data

	def getAddress(self):
		return (self.sheet, self.row, self.col)

	def getFormula(self):
		return self.data['formula']

	def getData(self):
		return [self.data, self.datatype]

	def getDataType(self):
		return self.datatype

	def calcOffset(self, relcell):
		col = self.col
		row = self.row
		if relcell[1]: col += relcell[1]
		if relcell[0]: row += relcell[0]
		return self.sheet, row, col

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
		if t[0] == 'FORMULA':
			cell.parsed = t[1]
			expand(t[1], cell)
		else:
			raise Exception("a formula must start with 'FORMULA', either the parser failed of it's an invalid cell")


def expand(tree, cell):
	'''start to drill down into the formula.'''
	print('Expanding: ', cell.getAddress())
	print(tree)
	for t in tree:
		if t[0] == 'FUNC':
			if t[1] == 'MAX':
				callStack.append(t[1])
				expand(t[2], cell)
			else:
				print("ERROR: "+aspect[1]+" function not written yet!")
				sys.exit(500)
		else:
			if t[0] == 'RELCELLRANGE':
				# fetchCellRange(t[1])
				# fetchCellRange(t[2])
				print("fetch data from range", cell.calcOffset(t[1]), cell.calcOffset(t[2]))
			elif t[0] == 'BINOP':
				args = expand(t[2], cell)
				print("binop", t[1], args)
			elif t[0] == 'RELCELL':
				# print("fetch data from cell", cell.calcOffset(t[1]))
				fetchCell(xlWB.getCell(cell.calcOffset(t[1])))
			else:
				print("ERROR: type", t, "not supported yet!")
				sys.exit(500)


def main(sourceFileName):

	global callStack, xlWB

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
	
	callStack = []
	fetchCell(cell)


if __name__ == '__main__':
	main('design_check.xml')
