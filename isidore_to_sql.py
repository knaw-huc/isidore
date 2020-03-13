# -*- coding: utf-8 -*-
import argparse
import csv
from datetime import datetime
import re
import sys
import xlrd 
import json
from pprint import pprint
import roman;
from roman import InvalidRomanNumeralError;

output = None
delimiter = ','
quotechar = ''


def xls_file(inputfiles, mapping, headerrow=0):
    for filename in inputfiles:
        wb = xlrd.open_workbook(filename,headerrow,encoding_override="utf-8") 
        sheet = wb.sheet_by_index(0) 
        result = []
        headers = []
        scaled_places = []
        has_scaled_place = {}
        includes_books = {}
        for colnum in range(sheet.ncols):
            if sheet.cell_value(headerrow,colnum) != '':
                headers.append(re.sub(r'[ -/]+','_',sheet.cell_value(headerrow,colnum)).strip('_'))
            else:
                headers.append("empty{}".format(colnum))
        teller = 0
        for rownum in range((headerrow+1), sheet.nrows):
            manuscript = []
            mid = "{}".format(sheet.cell_value(rownum,0))
            for colnum in range(sheet.ncols):
                cell_type = sheet.cell_type(rownum,colnum)
                cell = "{}".format(sheet.cell_value(rownum,colnum))
                # there are no dates in this file
#                if sheet.cell_type(rownum,colnum)==xlrd.XL_CELL_DATE:
#                    cell_date = xlrd.xldate.xldate_as_datetime(sheet.cell_value(rownum,colnum),0)
#                    stderr(cell_date.strftime("%d-%m-%Y"))
                if cell != '':
                    if cell_type==xlrd.XL_CELL_TEXT:
                        cell = re.sub(r'&','&amp;',cell)
                        cell = re.sub(r"\\","\\\\\\\\",cell)
                        manuscript.append(cell)
                    elif cell_type==xlrd.XL_CELL_NUMBER:
                        if cell.endswith('.0'):
                            cell = cell[0:-2]
                        if '.' in cell:
                            manuscript.append(cell)
                        else:
                            manuscript.append(f'{int(cell)}')
                    else:
                        stderr('Not found')
                    if headers[colnum]=="place_scaled":
                        if not cell in scaled_places:
                            scaled_places.append(cell)
                        has_scaled_place[mid] = cell
                    if headers[colnum]=="books_included":
                        res = try_roman(cell)
                        includes_books[mid] = res
                else:
                    manuscript.append('')
            result.append(manuscript)
            teller += 1

        output.write("DROP TABLE manuscripts cascade;\n")
        output.write("CREATE TABLE manuscripts (\n        ")
        all_headers = " text,\n        ".join(headers).replace("ID text","ID text primary key")
        output.write(all_headers)
        output.write(" text\n);\n")
        output.write("COPY manuscripts (")
        output.write(", ".join(headers))
        output.write(") FROM stdin;\n")
        for row in result:
            output.write("\t".join(row))
            output.write("\n")
        output.write("\\.\n")

        output.write("\n")
        output.write("DROP TABLE scaled_places cascade;\n")
        output.write("CREATE TABLE scaled_places (\n")
        output.write("        place text primary key\n");
        output.write(");\n")
        output.write("COPY scaled_places (place) FROM stdin;\n")
        scaled_places.sort()
        for place in scaled_places:
            output.write(f"{place}\n")
        output.write("\\.\n")

        output.write("\n")
        output.write("DROP TABLE manuscripts_scaled_places;\n")
        output.write("CREATE TABLE manuscripts_scaled_places (\n")
        output.write("        m_id text references manuscripts(ID),\n");
        output.write("        place text references scaled_places(place)\n");
        output.write(");\n")
        output.write("COPY manuscripts_scaled_places (m_id, place) FROM stdin;\n")
        for key in has_scaled_place.keys():
            output.write(f"{key}\t{has_scaled_place[key]}\n")
        output.write("\\.\n")

        output.write("\n")
        output.write("DROP TABLE books;\n")
        output.write("CREATE TABLE books(\n")
        output.write("        id int primary key,\n");
        output.write("        roman text\n");
        output.write(");\n")
        output.write("COPY books (id, roman) FROM stdin;\n")
        for i in range(1,21):
            output.write(f"{i}\t{roman.toRoman(i)}\n")
        output.write("\\.\n")

        output.write("\n")
        output.write("DROP TABLE manuscripts_books_included;\n")
        output.write("CREATE TABLE manuscripts_books_included (\n")
        output.write("        m_id text references manuscripts(ID),\n");
        output.write("        b_id int references books(id)\n");
        output.write(");\n")
        output.write("COPY manuscripts_books_included (m_id, b_id) FROM stdin;\n")
        for key in includes_books.keys():
            for book in includes_books[key]:
                output.write(f"{key}\t{book}\n")
        output.write("\\.\n")

        output.write("\n")

def try_roman(text):
#    stderr(text)
    try:
        n = roman.fromRoman(text)
#        stderr([n])
        return [n]
#        stderr(f'{text}: {n}')
    except InvalidRomanNumeralError:
        parts = re.split(r', *', text)
        if len(parts)>1:
            res = []
            for part in parts:
                res = res + try_roman(part)
            return res
        else:
            parts = re.split(r'-', text)
            if len(parts)==2:
                [first] = try_roman(parts[0])
                [last] = try_roman(parts[1])
                return list(range(first,last+1))
            else:
                stderr(f'{text} not valid?')


def clean_string(value):
    return value


def stderr(text):
    sys.stderr.write("{}\n".format(text))

def arguments():
    ap = argparse.ArgumentParser(description='Read file (csv, xls(x), sql_dump to make xml-rdf')
    ap.add_argument('-i', '--inputfile',
                    help="inputfile",
                    default = "20200227_manuscripts_mastersheet_CURRENT.xlsx")
    ap.add_argument('-m', '--mappingfile',
                    help="mappingfile (default = mapping.json)",
                    default = "mapping.json")
    ap.add_argument('-o', '--outputfile',
                    help="outputfile",
                    default = "isidore_test.sql")
    ap.add_argument('-q', '--quotechar',
                    help="quotechar",
                    default = "'" )
    ap.add_argument('-t', '--headerrow',
                    help="headerrow; 0=row 1 (default = 0)",
                    default = 0)
    args = vars(ap.parse_args())
    return args


def end_prog(code=0):
    stderr(datetime.today().strftime("%H:%M:%S"))
    stderr("einde")
    sys.exit(code)

 
if __name__ == "__main__":
    stderr("start")
    stderr(datetime.today().strftime("%H:%M:%S"))

    args = arguments()
    inputfiles = args['inputfile'].split(',')
    outputfile = args['outputfile']
    headerrow = args['headerrow']
    quotechar = args['quotechar']

    output = open(outputfile, "w", encoding="utf-8")
 
    '''
    mapping_etc = {}
    mapping = {}

    with open(args['mappingfile']) as f:
        mapping_etc = json.load(f)
        mapping = mapping_etc['mapping']
    '''
    xls_file(inputfiles, None) # mapping)
    

    end_prog(0)

