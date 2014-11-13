
from parse_xml import Cell, print_cells, offset_cell, \
                      process_shared_string_row, \
                      get_worksheets, get_shared_strings, parse_worksheet, \
                      post_process

from lxml import objectify
from lxml import etree
from nose import tools
from unittest import TestCase

class test_print_cells(TestCase):

    def setup(self, xml):
        cell = []
        parser = etree.XMLParser(ns_clean=True)
        tree = objectify.fromstring(xml,parser=parser)
        #print objectify.dump(cell)
        cell.append(Cell(tree))
        return cell

    def test_print_value_cell(self):
        xml = '''<c r="A3"><v>9.48</v></c>'''
        cell = self.setup(xml)
        string = print_cells(cell)
        tools.assert_equal(string, "va    A3    9.48")

    def test_print_formula_cell_value(self):
        '''
        Test correct formatting of a formula cell with numerical value
        '''
        xml = '''<c r="E7"><f ca="1">E6+$C$6</f><v>4.45</v></c>'''
        cell = self.setup(xml)
        string = print_cells(cell)
        tools.assert_equal(string, "fo    E7 = E6+$C$6    4.45")

    def test_print_formula_cell_text(self):
        '''
        Test correct formatting of a formula cell with string value
        '''
        xml = '''<c r="E7"><f ca="1">E6+$C$6</f><v>"This is a string"</v></c>'''
        cell = self.setup(xml)
        string = print_cells(cell)
        tools.assert_equal(string, "fo    E7 = E6+$C$6    This is a string")

class test_offset_cell(TestCase):

    def test_offset_cell(self):
        old_formula = "A1"
        offset = (0,0)
        new_formula = offset_cell(old_formula, offset)
        tools.assert_equal(old_formula,new_formula)

class test_process_shared_string_row(TestCase):

    def setup(self, xml):
        parser = etree.XMLParser(ns_clean=True)
        tree = objectify.fromstring(xml,parser=parser)
        return tree

    def test_formatted_cell(self):
        '''
        Test correct reading of formatted shared string
        '''
        xml = '''<si>
		<r>
			<t>Preparation for AR5</t>
		</r>
		<r>
			<rPr>
				<sz val="11"/>
				<color theme="1"/>
				<rFont val="Calibri"/>
				<family val="2"/>
				<scheme val="minor"/>
			</rPr>
			<t xml:space="preserve"> (see other file):</t>
		</r>
	</si>'''
        row = self.setup(xml)
        output = process_shared_string_row(row)
        tools.assert_equal(output, "Preparation for AR5 (see other file):")

    def test_plain_cell(self):
        '''
        Test correct reading of plain shared string
        '''
        xml = '''<si><t>Pig iron production</t></si>'''
        row = self.setup(xml)
        output = process_shared_string_row(row)
        tools.assert_equal(output, "Pig iron production")

    def test_small_formatted_cell(self):
        '''
        Test correct reading of small formatted shared string
        '''
        xml = '''<si><t>Non-metallic minerals (cem proxy)</t><phoneticPr fontId="29" type="noConversion"/></si>'''
        row = self.setup(xml)
        output = process_shared_string_row(row)
        tools.assert_equal(output, "Non-metallic minerals (cem proxy)")


class test_workbook(TestCase):

    def test(self):
        """
        Opens the test workbook and checks the print out matches the content
        """
        filename = "tests/test1.xlsx"
        sheets = list(get_worksheets(filename))
        string_dict = get_shared_strings("xl/sharedStrings.xml")

        sheet = sheets[0]
        output, shared_formulas = parse_worksheet(sheet, string_dict)
        post_process(output, shared_formulas, string_dict)
        tools.assert_equal(print_cells(output), "st    A1    Hello World")
