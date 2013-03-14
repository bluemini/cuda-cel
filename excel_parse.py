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
		self.xlData["fileName"] = ""

	def setFileName(self, fileName):
		self.xlData["fileName"] = fileName

	def getFileName(self):
		return self.xlData["fileName"]

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
		try:
			cell = ExcelCell(worksheet, 
				col, 
				row, 
				self.xlData["worksheets"][worksheet][col][row]
			)
		except Exception as e:
			if not worksheet in self.xlData["worksheets"]:
				print("ERROR:", worksheet, "worksheet is not found")
			elif not col in self.xlData["worksheets"][worksheet]:
				print("ERROR: column", col, "not found in workbook")
				print(self.xlData['worksheets'][worksheet])
			elif not row in self.xlData["worksheets"][worksheet][col]:
				print("ERROR: row", row, "not found in workbook")
			raise e
		return cell

	def getNamedCell(self, workbook, name):
		if name in self.xlData["namedCells"]:
			cellRange = self.xlData["namedCells"][name]
			if len(cellRange) == 1:
				print("named cell points to:", cellRange[0][0], cellRange[0][1], cellRange[0][2])
				return self.getCell( (cellRange[0][0], cellRange[0][2], cellRange[0][1]) )
		else:
			return None

	def setNamedCell(self, cellName, cell):
		if not cellName in self.xlData["namedCells"]:
			self.addNamedCell(cellName)
		addr = cell.getAddress()
		self.xlData["namedCells"][cellName].append((addr[0], addr[2], str(addr[1])))

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
		self.BigA = ord('A')
		self.sheet = sheet
		self.col = int(self.name2col(col))
		self.row = int(row)
		self.data = data

	def getAddress(self):
		return (self.sheet, str(self.row), str(self.col2Name(self.col) ) )

	def getFormula(self):
		if 'formula' in self.data:
			return self.data['formula']

	def getData(self):
		return self.data

	def getDataType(self):
		return self.datatype

	def getVar(self):
		return self.sheet + "__" + self.col2Name(self.col) + "__" + str(self.row)

	def calcOffset(self, relcell):
		'''A relcell always supplies offsets in int format, therefore we must first
		convert the Column to an int, before adding.'''
		row = str(int(self.row) + int(relcell[0]))
		col = self.col2Name(str(int(self.col) + int(relcell[1]) ) )
		return self.sheet, row, col

	def col2Name(self, columnValue):
		'''Converts an integer column value to an Excel Alpha based name.'''
		try:
			columnValue = int(columnValue)  # 26 should yield Z, which is A + 25
			x = math.floor(columnValue / 26) - 1
			y = columnValue % 26
			col = ""
			if x > 0:
				col = chr(self.BigA + x)
			return col + chr(self.BigA + y)
		except Exception as e:
			# print("col2Name", columnValue, e)
			pass
		return columnValue

	def name2col(self, columnValue):
		'''Converts a str column value to an integer.'''
		acc = 0
		try:
			for ch in columnValue:
				acc = (acc * 26) + ord(ch) - self.BigA
			return acc
		except Exception as e:
			# print("name2col", e)
			pass
		return columnValue



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
					self.wb.setCellData(ExcelCell(self.worksheet, self.col-1, self.row, None), 'formula', attrs['ss:Formula'])

			# finally if we have a data element, then look at it
			elif name == 'Data' and 'Cell' in self.state:
				if 'ss:Type' in attrs:
					self.wb.setCellData(ExcelCell(self.worksheet, self.col-1, self.row, None), 'datatype', attrs['ss:Type'])

			# save any named cell references
			elif name == 'NamedCell' and self.state[-2] == 'Cell':
				if 'ss:Name' in attrs:
					self.wb.setNamedCell(attrs['ss:Name'], ExcelCell(self.worksheet, self.col-1, self.row, None))
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
			cell = ExcelCell(self.worksheet, self.col-1, self.row, None)
			self.wb.setCellData(cell, 'content', content)

	def terminate(self):
		self.wb.dump()
		sys.exit()

	def getWorkbook(self):
		return self.wb


def fetchCell(cell):
	example_string = cell.getFormula()
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
			expanded = expand(t[1], cell)
			return expanded
		else:
			raise Exception("a formula must start with 'FORMULA', either the parser failed of it's an invalid cell")

	else:
		data = cell.getData()
		varName = cell.getVar()
		if not varName in declaredVars:
			print("CODE:", data['datatype'], cell.getVar(), '=', data['content'], ";")
			declaredVars.append(varName)
		return data


def expand(tree, cell):
	'''start to drill down into the formula. This is one large part of the 
	guts of the parser, taking elements and expanding their references and
	then parsing and expanding them, until we reach cells that do not need
	any further expansion...'''
	print('Expanding: ', cell.getAddress(), tree)
	for t in tree:
		if t[0] == 'FUNC':
			if t[1] == 'IF':
				callStack.append(t[1])
				if len(t[2]) != 3:
					errorMssg = "IF function requires exactly 3 arguments, " + str(len(t[2])) + " given"
					raise Exception(errorMssg)
				return expand(t[2], cell)

			if t[1] == 'IFERROR':
				callStack.append(t[1])
				if len(t[2]) != 2:
					errorMssg = "IFERROR function requires exactly 2 arguments, " + str(len(t[2])) + " given"
					raise Exception(errorMssg)
				return expand(t[2], cell)

			if t[1] == 'MAX':
				varss = expand(t[2], cell) # varss will return a list of cells
				celss = []
				for ci in range(len(varss)):
					celss.append( fetchCell(varss[ci]) )
				for ci in range(len(varss)):
					if ci == 0:
						print("CODE: tempMax = ", varss[ci].getVar(), ";")
					else:
						print("CODE: if (", varss[ci].getVar(), " > tempMax ) { tempMax =", varss[ci].getVar(), "; }")
				return 'tempMax'

			else:
				print("ERROR: "+t[1]+" function not written yet!")
				sys.exit(500)
		else:
			if t[0] == 'RELCELLRANGE':
				print("fetch data from range", cell.calcOffset(t[1]), cell.calcOffset(t[2]))
				rangeData = []
				fromCellRowOff	= t[1][0]
				fromCellColOff	= t[1][1]
				toCellRowOff 	= t[2][0]
				toCellColOff 	= t[2][1]
				# print(str(fromCellRowOff), str(fromCellColOff), str(toCellRowOff), str(toCellColOff))
				for r in range(fromCellRowOff, int(toCellRowOff)+1):
					for c in range(fromCellColOff, toCellColOff+1):
						# print("fetch data from", cell.calcOffset( (r, c) ))
						rangeData.append(xlWB.getCell(cell.calcOffset( (r, c) ) ) )
				return rangeData

			elif t[0] == 'BINOP':
				args = expand(t[2], cell)
				print("binop", t[1], args)

			elif t[0] == 'RELCELL':
				# print("fetch data from cell", cell.calcOffset(t[1]))
				return fetchCell(xlWB.getCell(cell.calcOffset(t[1])))

			elif t[0] == 'NAME':
				# first check if this is a named cell reference
				namedCell = xlWB.getNamedCell(None, t[1])
				if namedCell:
					print("we have a named cell, so expand it..")
					fetchCell(namedCell)
				cell = xlWB.getCell(cell.calcOffset(t[1]))
				print("CODE: ", cell.getVar(), "=", cell.getData())
				return fetchCell()

			elif t[0] == 'NUMBER':
				print("We have a NUMBER value!!", t[1])
				return t

			elif t[0] == 'STRING':
				print("We have a value!!", t[1])
				return t

			elif t[0] == 'SUBEXP':
				print("Expanding sub expression")
				ev = expand(t[1], cell)
				return ev

			else:
				print("ERROR: type", t, "not supported yet!")
				sys.exit(500)


def main(sourceFileName, refresh=False):

	global callStack, xlWB, declaredVars

	declaredVars = []

	try:
		# if we are refreshing, raise an exception immediately
		if refresh:
			raise Exception("Refresh file")
		xlpickle = open("xlstruct.pkl", "r+b")
		U = pickle.Unpickler(xlpickle)
		xlWB = U.load()
		# finally check the name of the WB
		if xlWB.getFileName() != sourceFileName:
			raise Exception("Different file in pickle jar, must rerun parser")

	except:
		xlpickle = open("xlstruct.pkl", "w+b")
		source = open(sourceFileName, 'r', encoding='utf-8')
		handler = ExcelHandler()
		xml.sax.parse(source, handler)

		xlWB = handler.getWorkbook()
		xlWB.setFileName(sourceFileName)
		P = pickle.Pickler(xlpickle)
		P.dump(xlWB)

	# example_string = xlWB.getCell('PegTop', 'ES', '252', 'formula')
	cell = xlWB.getNamedCell(None, 'YResult')
	# example_string = "=IF(R[-277]C>R[-305]C-2*R[-303]C,'-',0.58*R[-319]C*PI()*R[-277]C*R[-303]C*R[-3]C/R[-404]C/1000)"
	
	callStack = []

	cell = xlWB.getCell(("Sheet1", "2", "D"))
	comp = fetchCell(cell)
	print(comp)


if __name__ == '__main__':
	# main('design_check.xml')
	main('stage2.xml', refresh=True)
