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
datasetname = ''
resource = ''


def xls_file(inputfiles, mapping, headerrow=0):
    for filename in inputfiles:
        wb = xlrd.open_workbook(filename,headerrow,encoding_override="utf-8") 
        sheet = wb.sheet_by_index(0) 
        result = []
        headers = []
        for colnum in range(sheet.ncols):
            if sheet.cell_value(headerrow,colnum) != '':
                headers.append(re.sub(r'[ -/]+','_',sheet.cell_value(headerrow,colnum)).strip('_'))
            else:
                headers.append("empty{}".format(colnum))
        teller = 0
        for rownum in range((headerrow+1), sheet.nrows):
            manuscript = {}
            for colnum in range(sheet.ncols):
                cell_type = sheet.cell_type(rownum,colnum)
                cell = "{}".format(sheet.cell_value(rownum,colnum))
#                if teller<5 and colnum==2:
#                    stderr("{} - {} - {}".format(sheet.cell_type(rownum,colnum),sheet.cell_value(rownum,colnum), cell_type))
                if sheet.cell_type(rownum,colnum)==xlrd.XL_CELL_DATE:
                    cell_date = xlrd.xldate.xldate_as_datetime(sheet.cell_value(rownum,colnum),0)
                    stderr(cell_date.strftime("%d-%m-%Y"))
                    pass
                if cell != '':
                    if cell_type==xlrd.XL_CELL_TEXT:
                        cell = re.sub(r'&','&amp;',cell)
                        manuscript[headers[colnum]] = cell
                    elif cell_type==xlrd.XL_CELL_NUMBER:
                        if cell.endswith('.0'):
                            cell = cell[0:-2]
                        manuscript[headers[colnum]] = cell if '.' in cell else int(cell)
                    else:
                        stderr('Not found')
                    if headers[colnum]=="books_included":
                        res = try_roman(cell)
                        manuscript['books_included_numerical'] = res
            result.append(manuscript)
            teller += 1
        output.write(json.dumps(result,sort_keys=False,indent=2, ensure_ascii=False))

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
                    default = "isidore_test.json")
    ap.add_argument('-r', '--resource',
                    help="resource",
                    default = "https://resource.huygens.knaw.nl/isidore/" )
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
    resource = args['resource']
    headerrow = args['headerrow']
    quotechar = args['quotechar']
    datasetname = re.search(r'/([^/]+)/$',resource).group(1)

    output = open(outputfile, "w", encoding="utf-8")
    
    mapping_etc = {}
    mapping = {}

    with open(args['mappingfile']) as f:
        mapping_etc = json.load(f)
        mapping = mapping_etc['mapping']

    xls_file(inputfiles, mapping)


    end_prog(0)

