# -*- coding: utf-8 -*-
from lxml import etree
from lxml import objectify

from excelutil import *
import tokenizer

class Cell(object):

    def __init__(self, address):
        self.address = address
        self.cell_type=None
        self.value=None
        self.formula=None
        self.formula_type=None
        self.formula_host=False
        self.formula_range=None
        self.shared_index=None

    def set_cell_type(self, cell_type):
        self.cell_type = cell_type

    def set_cell_value(self, cell_value):
        self.value = cell_value

    def set_formula_type(self,formula_type):
        self.formula_type = formula_type

    def set_formula_host(self,formula_host):
        self.formula_host = formula_host

    def set_formula(self, formula):
        self.formula = formula

    def set_formula_range(self, formula_range):
        self.formula_range = formula_range

    def set_shared_index(self,shared_index):
        self.shared_index=shared_index

    def set_value(self, value):
        self.value = value

    def pretty_print(self):
        if self.cell_type == "string":
            print "{:>3} \t {:<10}  \t {}".format(self.address,"", int(self.value))
        elif self.cell_type == "value":
            print "{:>3} \t {:<10}\t {}".format(self.address,"", round(float(self.value),2))
        else:
            print "{:>3} \t ={:<10} \t {}".format(self.address, self.formula, round(float(self.value),2))

    def debug_print(self):
        if self.cell_type == "formula":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)
        elif self.cell_type == "array":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)
        elif self.cell_type == "shared":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)
        elif self.cell_type == "string":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)
        elif self.cell_type == "value":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)

class Row(object):

    def __init__(self,row_num,span):
        self.row_num = row_num
        self.span = span
        self.cells = []

    def add_cell(self, address):
        self.cells.append(Cell(address))

    def pretty_print(self):
        for row in self.row_num:
            print "Row: {}".format(row)
        for cell in self.cells:
            cell.pretty_print()
#
#    def __iter__(self):
#        return self
#    def next(self):
#        if self.index == 0:
#            raise StopIteration
#        self.index = self.index - 1
#        return self.data[self.index]

def split_address(address):

    #ignore case
    address = address.upper()

    # regular <col><row> format
    if re.match('^[\$A-Z]+[\$\d]', address):
        col,row = filter(None,re.split('([\$]?[A-Z]+)',address))
    else:
        raise Exception('Invalid address format ' + address)

    return (col,row)

def check_address(address):
    if (address.find("$") != -1):
        # address has an absolute
        return True
    else:
        return False

def compute_offset(host_address, client_address):
    '''
    Returns the absolute difference between two addresses
    '''
    host_col, host_row = address2index(host_address)
    client_col, client_row = address2index(client_address)

    column_offset = client_col-host_col
    row_offset = client_row-host_row

    return tuple((column_offset,row_offset))


file = "xl/worksheets/sheet1.xml"
parser = etree.XMLParser(ns_clean=True)
tree = objectify.parse(file, parser)
root = tree.getroot()
ns = root.nsmap

output = []

shared_formulas = []

rows = list(root)[3]
for row in rows:

    row_num = row.attrib.get("r")
    span = row.attrib.get("span")
    #output.append(Row(row_num, span))
    #print "\nRow: {}".format(row_num)

    cells = list(row)

    for cell in cells:

        cell_address = ""
        cell_type = ""
        cell_value = ""
        cell_shared_host = False
        cell_formula = ""
        cell_formula_type = ""
        cell_formula_range = ""
        cell_shared_index = None

        tags = []

        items = list(cell)

        cell_address = cell.attrib.get("r")

        # Add a cell to the list of cells
        output.append(Cell(cell_address))

        # Get a tempory list of the tags
        for item in items:
            tags.append(item.tag[-1])

        if cell.attrib.get("t") == "s":
            cell_type = "string"
            output[-1].set_cell_type(cell_type)
            for item in items:
                if item.tag[-1] == "v":
                    cell_value = item.text # lookup to string table via cell.text
                    output[-1].set_cell_value(cell_value)
        elif (not "f" in tags): # look to see if there is a formula - if not, it is a value
            cell_type = "value"
            output[-1].set_cell_type(cell_type)
            cell_value = item.text
            output[-1].set_cell_value(cell_value)
        else:
            for item in items:
                #print "Cell {}:".format(cell_address)
                #print "Attributes {}".format(item.attrib)
                if item.tag[-1] == "f":
                    if item.attrib.get("t") == "array":
                        cell_type = "array"
                        output[-1].set_cell_type(cell_type)
                        cell_formula = item.text
                        output[-1].set_formula(cell_formula)
                        #print "Array formula"
                    elif item.attrib.get("t") == "shared":
                        cell_type = "shared"
                        output[-1].set_cell_type(cell_type)
                        #print "Shared formula: {} {}".format(item.attrib,item.text)
                        if item.attrib.get("ref"):
                            #print "Host cell of si {}".format(item.attrib.get("si"))
                            cell_shared_host = True
                            output[-1].set_formula_host(cell_shared_host)
                            cell_shared_index = int(item.attrib.get("si"))
                            output[-1].set_shared_index(cell_shared_index)
                            cell_formula = item.text
                            output[-1].set_formula(cell_formula)
                            shared_formulas.append(dict(si=int(cell_shared_index),formula=cell_formula,address=cell_address))
                        else:
                            cell_shared_index = int(item.attrib.get("si"))
                            output[-1].set_shared_index(cell_shared_index)
                            cell_formula = "si {}".format(cell_shared_index)
                            output[-1].set_formula(cell_formula)
                    else:
                        cell_type = "formula"
                        output[-1].set_cell_type(cell_type)
                        cell_formula = item.text
                        output[-1].set_formula(cell_formula)

                    #print "Item {}: and type {} \t {}".format(item.tag[-1] , "formula" , item.text)
                elif item.tag[-1] == "v":
                    cell_value = item.text
                    output[-1].set_cell_value(cell_value)
                    #print "Item {}: and type {} \t {}".format(item.tag[-1] , "value" , item.text)
        #print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(cell_address, cell_type, cell_formula, round(float(cell_value),2), cell_shared_host, cell_shared_index)
#print ""
#for cell in output:
#    cell.pretty_print()
#next((item for item in dicts if item["name"] == "Pam"), None)

for cell in output:
    new_formula = []
    if (cell.cell_type == "shared") & (cell.formula_host == False):
        # Cell is a shared formula
        cell.shared_index
        expression = next((formula["formula"] for formula in shared_formulas if formula["si"] == cell.shared_index),None)
        host_address = next((formula["address"] for formula in shared_formulas if formula["si"] == cell.shared_index),None)
        client_address = cell.address

        p = tokenizer.ExcelParser()
        p.parse(expression)

        offset = compute_offset(host_address, client_address)

        for t in p.tokens.items:
            if t.ttype == "operand" and t.tsubtype == "range":
                if check_address(t.tvalue) == False:
                    formula_range = address2index(t.tvalue)
                    col,row = map(sum,zip(formula_range,offset))
                    new_formula.append(index2addres(col,row))
                else:
                    col, row = split_address(t.tvalue)
                    if check_address(col) == False:
                        # Column is not absolute address
                        colnum = col2num(col)
                        col = num2col(colnum + offset[0])
                    if check_address(row) == False:
                        # Row is not absolute address
                        row = row + offset[1]
                    new_formula.append("".join([col,row]))
            else:
                new_formula.append(t.tvalue)
        cell.set_formula(''.join(new_formula))

for cell in output:
    #cell.debug_print()
    cell.pretty_print()
