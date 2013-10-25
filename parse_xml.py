# -*- coding: utf-8 -*-
from lxml import etree
from lxml import objectify
import zipfile
import re
import sys

from excelutil import col2num, num2col, address2index, index2addres
from tokenizer import ExcelParser

shared_formulas = []

class Cell(object):

    def __init__(self, cell):
        self.address = cell.attrib.get("r")
        self.cell_type = None
        self.value = None
        self.formula = None
        self.formula_type = None
        self.formula_host = False
        self.formula_range = None
        self.shared_index = None

        items = list(cell)
        tags = []
        # Get a tempory list of the tags
        for item in items:
            tags.append(item.tag[-1])

        if cell.attrib.get("t") == "s": # cell is of type string
            cell_type = "string"
            self.set_cell_type(cell_type)
            for item in items:
                if item.tag[-1] == "v":
                    cell_value = str(item.text) # lookup to string table via cell.text
                    self.set_cell_value(cell_value)
        elif (not "f" in tags): # look to see if there is a formula - if not, it is a value
            cell_type = "value"
            self.set_cell_type(cell_type)
            for item in items:
                cell_value = float(item.text)
                self.set_cell_value(cell_value)
        else: # otherwise it is an array/shared/formula cell
            for item in items: # Iterate over the attributes of the cell
                if item.tag[-1] == "f":
                    if item.attrib.get("t") == "array":
                        cell_type = "array"
                        self.set_cell_type(cell_type)
                        cell_formula = item.text
                        self.set_formula(cell_formula)
                    elif item.attrib.get("t") == "shared":
                        cell_type = "shared"
                        self.set_cell_type(cell_type)
                        if item.attrib.get("ref"):
                            cell_shared_host = True
                            self.set_formula_host(cell_shared_host)
                            cell_shared_index = int(item.attrib.get("si"))
                            self.set_shared_index(cell_shared_index)
                            cell_formula = item.text
                            self.set_formula(cell_formula)
                            global shared_formulas
                            shared_formulas.append(dict(si=int(cell_shared_index),formula=cell_formula,address=self.address))
                        else:
                            cell_shared_index = int(item.attrib.get("si"))
                            self.set_shared_index(cell_shared_index)
                            cell_formula = "si {}".format(cell_shared_index)
                            self.set_formula(cell_formula)
                    else:
                        cell_type = "formula"
                        self.set_cell_type(cell_type)
                        cell_formula = item.text
                        self.set_formula(cell_formula)
                elif item.tag[-1] == "v":
                    cell_value = item.text
                    self.set_cell_value(cell_value)

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
            print '%(address)06s %(value)-20s' % \
                {"address" : self.address, "value" : self.value }
        elif self.cell_type == "value":
            print '%(address)06s %(value) 60.2f' % \
                {"address" : self.address, "value" : float(self.value) }
        else:
            print '%(address)06s =%(formula)-20s %(value) 38.2f' % \
                {"address" : self.address, "value" : float(self.value), "formula" : self.formula }

    def debug_print(self):
        if self.cell_type == "formula":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, self.value, self.formula_host, self.shared_index)
        elif self.cell_type == "array":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, self.value, self.formula_host, self.shared_index)
        elif self.cell_type == "shared":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, self.value, self.formula_host, self.shared_index)
        elif self.cell_type == "string":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, self.value, self.formula_host, self.shared_index)
        elif self.cell_type == "value":
            print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(self.address, self.cell_type, self.formula, round(float(self.value),2), self.formula_host, self.shared_index)

def split_address(address):

    #ignore case
    address = address.upper()

    # regular <col><row> format
    if re.match('^[\$A-Z]+[\$\d]', address):
        col,row = filter(None,re.split('([\$]?[A-Z]+)',address))
    else:
        raise Exception('Invalid address format ' + address)

    return (col,row)

def check_address(address,symbol):
    if (address.find(symbol) != -1):
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

def get_worksheets(name):
   arc= zipfile.ZipFile( name, "r" )
   member= arc.getinfo("xl/sharedStrings.xml")
   arc.extract( member )
   for member in arc.infolist():
       if member.filename.startswith("xl/worksheets") and member.filename.endswith('.xml'):
           arc.extract(member)
           yield member.filename

def get_shared_strings(shared_strings_file):

    shared_string_dict = []
    parser = etree.XMLParser(ns_clean=True)
    stree = objectify.parse(shared_strings_file, parser)
    sroot = stree.getroot()
    srows = list(sroot)
    for index, srow in enumerate(srows):
        shared_string_dict.append(srow[0].text)
    return shared_string_dict

def get_row(row_name, tree_root):
    return next((row for row in list(tree_root) if row.tag == "{" + tree_root.nsmap.get(None) + "}"+ row_name), None)

def parse_worksheet(sheetname, string_dict):
    '''
    Returns:
        a list of class Cells
        a list of shared_formulas
    '''
    parser = etree.XMLParser(ns_clean=True)
    tree = objectify.parse(sheetname, parser)
    root = tree.getroot()

    # A list of cells
    output = []

    # A list of shared formulas
    global shared_formulas
    shared_formulas = []

    rows = get_row("sheetData", root)
    for row in rows: # Iterate over the rows

        cells = list(row)

        for cell in cells: # Iterate over the cells in a row

            # Add a cell to the list of cells
            output.append(Cell(cell))

    return output, shared_formulas

def post_process(output, shared_formulas, string_dict):
    '''

    '''
    for cell in output:
        #print cell.address, cell.formula, cell.cell_type, cell.formula_host
        if (cell.cell_type == "shared") & (cell.formula_host == False):
            cell.set_formula(update_shared_formulas(cell, shared_formulas))
            #except:
                #print "ERROR", cell.address
        elif (cell.cell_type == "string"):
            cell.set_value( string_dict[int(cell.value)])

def update_shared_formulas(cell, shared_formulas):

    new_formula = []

    expression = next((formula["formula"] for formula in shared_formulas if formula["si"] == cell.shared_index),None)
    host_address = next((formula["address"] for formula in shared_formulas if formula["si"] == cell.shared_index),None)
    client_address = cell.address

    p = ExcelParser()
    p.parse(expression)

    offset = compute_offset(host_address, client_address)

    for t in p.tokens.items: # Iterate over the tokens
        if t.ttype == "operand" and t.tsubtype == "range":
            if check_address(t.tvalue,":") == True: # If operand-range is a range
                # split the range
                new_range = []
                for ad in t.tvalue.split(":"):
                    updated_address = offset_cell(ad,offset)
                    new_range.append(updated_address)
                new_formula.append(":".join(new_range))
            else: # If operand-range is just a cell
                updated_address = offset_cell(t.tvalue,offset)
                new_formula.append(updated_address)
        else: # If not a range
            new_formula.append(t.tvalue)

    return ''.join(new_formula)

def offset_cell(address,offset):
    new_formula = []
    if check_address(address,"$") == False:
        formula_range = address2index(address)
        col,row = map(sum,zip(formula_range,offset))
        new_formula = index2addres(col,row)
    else:
        col, row = split_address(address)
        if check_address(col,"$") == False:
            # Column is not absolute address
            colnum = col2num(col)
            col = num2col(colnum + offset[0])
        if check_address(row,"$") == False:
            # Row is not absolute address
            row = row + offset[1]
        new_formula = "".join([col,row])
    return new_formula

def print_cells(output):
    for cell in output:
        #cell.debug_print()
        cell.pretty_print()

def main():
    args = sys.argv[1:]
    if len(args) != 1:
        print 'usage: python parse_xml.py infile.xlsx'
        sys.exit(-1)
    #outfile = sys.stdout
    sheets = list(get_worksheets(args[0]))
    string_dict = get_shared_strings("xl/sharedStrings.xml")
    for sheet in sheets:
        print sheet
        output, shared_formulas = parse_worksheet(sheet,string_dict)
        post_process(output, shared_formulas, string_dict)
        print_cells(output)

if __name__ == '__main__':
    main()
    #pass

#filename = "ipcc.xlsx"
#sheets = list(get_worksheets(filename))
#string_dict = get_shared_strings("xl/sharedStrings.xml")
#
#sheet = sheets[6]
##for sheet in sheets:
#print sheet
#output, shared_formulas = parse_worksheet(sheet,string_dict)
#post_process(output, shared_formulas, string_dict)
#print_cells(output)
