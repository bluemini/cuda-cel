import xml.sax
import sys
import math
import excel_lang
import pickle

class ExcelHandler(xml.sax.ContentHandler):

	def __init__(self):
		xml.sax.ContentHandler.__init__(self)
		self.ok = False
		self.state = []

		self.xlData = {}
		self.xlData["namedRanges"] = []
		self.xlData["namedCells"]  = {}
		self.xlData["worksheets"]  = {}

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
				self.xlData["namedRanges"].append({'name':attrs['ss:Name'], 'refersto':attrs['ss:RefersTo']})
				# print("Named Range: " + str(self.xlData["namedRanges"][len(self.xlData["namedRanges"])-1]))

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
					self.setCellData(self.worksheet, self.col, self.row, 'formula', attrs['ss:Formula'])

			# finally if we have a data element, then look at it
			elif name == 'Data' and 'Cell' in self.state:
				# print("got data")
				pass

			# save any named cell references
			elif name == 'NamedCell' and self.state[-2] == 'Cell':
				if 'ss:Name' in attrs:
					if not attrs['ss:Name'] in self.xlData["namedCells"]:
						self.xlData["namedCells"][attrs['ss:Name']] = []
					self.xlData["namedCells"][attrs['ss:Name']].append((self.col2Name(self.col), str(self.row)))
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
			self.setCellData(self.worksheet, self.col, self.row, 'content', content)

	def col2Name(self, columnValue):
		columnValue -= 1  # 26 should yield Z, which is A + 25
		x = math.floor(columnValue / 26) - 1
		y = columnValue % 26
		BigA = ord('A')
		col = ""
		if x > 0:
			col = chr(BigA + x)
		return col + chr(BigA + y)

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

	def terminate(self):
		print(self.xlData["worksheets"])
		print(self.xlData["namedCells"])
		print(self.xlData["namedRanges"])
		sys.exit()

	def getWorksheets(self):
		return self.xlData


def main(sourceFileName):

	try:
		xlpickle = open("xlstruct.pkl", "r+b")
		U = pickle.Unpickler(xlpickle)
		xlData = U.load()

	except:
		xlpickle = open("xlstruct.pkl", "w+b")
		source = open(sourceFileName, 'r', encoding='utf-8')
		handler = ExcelHandler()
		xml.sax.parse(source, handler)

		xlData = handler.getWorksheets()
		P = pickle.Pickler(xlpickle)
		P.dump(xlData)

	example_string = xlData["worksheets"]['PegTop']['ES']['252']["formula"]

	# Give the lexer some input
	# example_string = "=IF(R[-277]C>R[-305]C-2*R[-303]C,'-',0.58*R[-319]C*PI()*R[-277]C*R[-303]C*R[-3]C/R[-404]C/1000)"
	print(example_string)
	excel_lang.lexer.input(example_string)

	# Tokenize
	while True:
	    tok = excel_lang.lexer.token()
	    if not tok: break      # No more input
	    print(tok)

	t = excel_lang.yacc.parse(example_string)

if __name__ == '__main__':
	main('design_check.xml')
