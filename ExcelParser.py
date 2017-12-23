#dependency: openpyxl

import string

import openpyxl

class ExcelParser:
    treshhold = 5
    
    def __init__(self, file, output="database_dump"):
        self.document = openpyxl.load_workbook(file)
	self.output = output

    def sheets(self):
        sheets = self.document.get_sheet_names()
        return sheets

    def extract(self, sheet):
	data = []
	rows = list(self.document.get_sheet_by_name(sheet).rows)
	
	row_miss= 0
	for index, row in enumerate(rows):
	    if row[0].value is None:
		row_miss +=1
#		continue

	    if row_miss > self.treshhold:
		break; #eof
	    record = []
	    col_miss = 0
	    for cell in row:
		if cell.value is None:
		    col_miss += 1
		    #continue

		if col_miss > self.treshhold:
		    break

		record.append(cell.value)

	    if (record[0] is None): #asume merge
		while (len(record) > len(data[-1])):
		    data[-1].append(None)
		for index, update in reversed(list(enumerate(record))):
		    if update is not None:
			try:
			    if data[-1][index] is None:	
				data[-1][index] = temp[index-1] + " " + update
			    else:
				data[-1][index] += " "+ str(update)
			except Exception as e:
			    pass
		
	    else:
	        while record[-1] is None:
		    del record[-1]

	        data.append(record)

	return data

    def out(self):
	for sheet in self.sheets():
	    self.parse(self.extract(sheet), sheet)
	print "SQL DATA HAVE BEEN GENERATED"
    def parse(self, data, sheet):
	description = ""
	bin = []
	for index, i in enumerate(data):
	    if len(i)==1:
		description += "\n-- " + str(i[0])
		bin.append(index)
	table_name = self.to_usable(sheet)
	for b in reversed(bin):
	    del data[b]
	size = len(data[0])
        while data[0][-1] is None and len(data[1]) != size:
            del data[0][-1]
            size =len(data[0])
	try:
   	    for row in data:
	        if len(row) != size:
		    raise Exception("Data integritiy could not be established, there are rows with varying columns on {} at {}".format(description, row))
	except Exception as e:
	    pass
	sql = description
	sql += "\n CREATE DATABASE IF NOT EXISTS {0} CHARSET=UTF8; USE {0};".format(self.output)
	sql += "\n--Schema"
	sql += "\nCREATE TABLE IF NOT EXISTS `{}`(".format(table_name)

	for c, column_name in enumerate(data[0]):
	    if column_name is None:
		column_name = data[0][c-1] + "_2"
	    sql += "`{}` {},".format(self.to_usable(column_name), self.sql_type(data[1:], c))
	sql = sql.rstrip(',')
	sql += ");"
	sql += "-- data"
	for row in data[1:]:
	    sql += "\nINSERT INTO `{}` VALUES('{}');".format(table_name, "','".join(map(str, row)))
        self.write_out(sql)

    def to_usable(self, s):
	s = str(s)
	s.replace(".", "_")
	return s.replace(" ", "_")

    def sql_type(self, data, position):
	_type = ""
	
	for row in data:
	    if type(row[position]) is int and _type != "VARCHAR(255)":
		_type = "INT"
	    elif type(row[position]) is float and _type != "VARCHAR(255)":
		_type =  "DECIMAL (20,2)"
	    else:
               _type = "VARCHAR(255)" 
	return _type

    def write_out(self, sql):
	with open (self.output, 'a') as file:
	    file.write(sql)

