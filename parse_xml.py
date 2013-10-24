# -*- coding: utf-8 -*-
from lxml import etree
from lxml import objectify

from excelutil import *
import tokenizer

class Cell(object):

    def __init__(self):
        self.attributes = dict(address=None,cell_type=None,value=None, formula=None, \
            formula_type=None,formula_host=False,formula_range=None)

    def add_cell(self, address, cell_type):
        self.attributes["address"] = address
        self.attributes["cell_type"] = cell_type

    def set_formula_type(self,formula_type):
        self.formula_type = formula_type

    def set_formula_host(self,formula_host):
        self.formula_host = formula_host

    def set_formula(self, formula):
        self.formula = formula

    def set_formula_range(self, formula_range):
        self.formula_range = formula_range

    def set_value(self, value):
        self.value = value

    def pretty_print(self):
        address = self.attributes.get("address")
        value = self.attributes.get("value")
        cell_type = self.attributes.get("cell_type")
        if cell_type == "formula":
            formula = self.attributes.get("formula")
            print "{} \t : {} \t : {}".format(address, formula, value)
        elif cell_type == "string":
            print "{} \t : {}".format(address, value)

class Row(object):

    def __init__(self):
        self.attributes = dict(row_num=None,span=None,cells=Cell())

    def add_row(self,row_num,span):
        self.attributes["row_num"] = row_num
        self.attributes["span"] = span

    def add_cell(self, address, cell_type):
        self.attributes["cells"].add_cell(address,cell_type)

    def pretty_print(self):
        for row in self.attributes.get("row_num"):
            print "Row: {}".format(row)

    def __iter__(self):
        return self
    def next(self):
        if self.index == 0:
            raise StopIteration
        self.index = self.index - 1
        return self.data[self.index]

file = "xl/worksheets/sheet1.xml"
parser = etree.XMLParser(ns_clean=True)
tree = objectify.parse(file, parser)
root = tree.getroot()
ns = root.nsmap

output = Row()

shared_formulas = []

rows = list(root)[3]
for row in rows:

    row_num = row.attrib.get("r")
    span = row.attrib.get("span")
    #output.add_row(row_num, span)
    print "\nRow: {}".format(row_num)

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
        for item in items:
            tags.append(item.tag[-1])
        if cell.attrib.get("t") == "s":
            cell_type = "string"
            output.add_cell(cell_address, "string")
            for item in items:
                if item.tag[-1] == "v":
                    cell_value = item.text # lookup to string table via cell.text
        elif (not "f" in tags): # look to see if there is a formula - if not, it is a value
            cell_type = "value"
            cell_value = item.text
        else:
            for item in items:
                #print "Cell {}:".format(cell_address)
                #print "Attributes {}".format(item.attrib)
                if item.tag[-1] == "f":
                    if item.attrib.get("t") == "array":
                        cell_type = "array"
                        cell_formula = item.text
                        #print "Array formula"
                    elif item.attrib.get("t") == "shared":
                        cell_type = "shared"
                        #print "Shared formula: {} {}".format(item.attrib,item.text)
                        if item.attrib.get("ref"):
                            #print "Host cell of si {}".format(item.attrib.get("si"))
                            cell_shared_host = True
                            cell_shared_index = item.attrib.get("si")
                            cell_formula = item.text
                            shared_formulas.append(dict(si=cell_shared_index,formula=cell_formula))
                        else:
                            cell_shared_index = item.attrib.get("si")
                            cell_formula = "si {}".format(cell_shared_index)
                    else:
                        cell_type = "formula"
                        cell_formula = item.text
                        output.add_cell(cell_address,"formula")
                    #print "Item {}: and type {} \t {}".format(item.tag[-1] , "formula" , item.text)
                elif item.tag[-1] == "v":
                    cell_value = item.text
                    #print "Item {}: and type {} \t {}".format(item.tag[-1] , "value" , item.text)
        print "Cell {:>3} is a {} \t {:<10} \t {} \t host:{} \t si:{}".format(cell_address, cell_type, cell_formula, round(float(cell_value),2), cell_shared_host, cell_shared_index)
#for op in list(output):
#    op.pretty_print()
#next((item for item in dicts if item["name"] == "Pam"), None)
