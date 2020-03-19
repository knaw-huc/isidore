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
schema_out = None
delimiter = ','
quotechar = ''
pattern = re.compile(r'([^(]*)\(([^)]*)\)([^(]*)')


def xls_file(inputfiles, headerrow=0):
    for filename in inputfiles:
        # after converting de xlsx sheet to xls, the color information can be extracted and used
        # not sure if we will do this
        if filename.endswith("xls"):
            xls_type = True
        else:
            xls_type = False
        if xls_type:
            wb = xlrd.open_workbook(filename,headerrow,encoding_override="utf-8", formatting_info=True)
        else:
            wb = xlrd.open_workbook(filename,headerrow,encoding_override="utf-8")
        sheet = wb.sheet_by_index(0) 
        result = []
        headers = []
        scaled_places = []
        has_scaled_place = {}
        scaled_dates = []
        has_scaled_date = {}
        content_types = []
        has_content_type = {}
        includes_books = {}
        location_details = []

        for colnum in range(sheet.ncols):
            if sheet.cell_value(headerrow,colnum) != '':
                headers.append(re.sub(r'[ -/]+','_',sheet.cell_value(headerrow,colnum)).strip('_'))
            else:
                headers.append("empty{}".format(colnum))
        teller = 0
        for rownum in range((headerrow+1), sheet.nrows):
            manuscript = []
            m_id = "{}".format(sheet.cell_value(rownum,0))
            for colnum in range(sheet.ncols):
                cell_type = sheet.cell_type(rownum,colnum)
                cell = "{}".format(sheet.cell_value(rownum,colnum))
                if xls_type:
                    color = getBGColor(wb, sheet, rownum, colnum)
                    if color:
                        # will see if we will do something with the color information
                        # stderr(f'{cell} ({rownum},{colnum}): {color}')
                        pass
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
                        has_scaled_place[m_id] = cell
                        manuscript.pop()
                    elif headers[colnum]=="date_scaled":
                        if not cell in scaled_dates:
                            scaled_dates.append(cell)
                        has_scaled_date[m_id] = cell
                        manuscript.pop()
                    elif headers[colnum]=="books_included":
                        res = try_roman(cell)
                        includes_books[m_id] = res
                        manuscript.pop()
                    elif headers[colnum]=="content_type":
                        if not cell in content_types:
                            content_types.append(cell)
                        has_content_type[m_id] = cell
                        manuscript.pop()
                elif not headers[colnum] in ["place_scaled","date_scaled","books_included","content_type"]:
                    manuscript.append('')
            handle_content_detail(location_details,m_id, sheet.cell_value(rownum,27),sheet.cell_value(rownum,28))
            result.append(manuscript)
            teller += 1

        headers_2 = []
        for header in headers:
            if not header in ["place_scaled","date_scaled","books_included","content_type"]:
                headers_2.append(header + " text")
        headers_2[0] += " primary key"
        create_schema(headers_2)

        headers_2 = []
        for header in headers:
            if not header in ["place_scaled","date_scaled","books_included","content_type"]:
                headers_2.append(header)
        output.write("COPY manuscripts (")
        output.write(", ".join(headers_2))
        output.write(") FROM stdin;\n")
        for row in result:
            output.write("\t".join(row))
            output.write("\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_places (place) FROM stdin;\n")
        scaled_places.sort()
        for place in scaled_places:
            output.write(f"{place}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scaled_places (m_id, place) FROM stdin;\n")
        for key in has_scaled_place.keys():
            output.write(f"{key}\t{has_scaled_place[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_dates (date) FROM stdin;\n")
        scaled_dates.sort()
        for date in scaled_dates:
            output.write(f"{date}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scaled_dates (m_id, date) FROM stdin;\n")
        for key in has_scaled_date.keys():
            output.write(f"{key}\t{has_scaled_date[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY books (id, roman) FROM stdin;\n")
        for i in range(1,21):
            output.write(f"{i}\t{roman.toRoman(i)}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_books_included (m_id, b_id) FROM stdin;\n")
        for key in includes_books.keys():
            for book in includes_books[key]:
                output.write(f"{key}\t{book}\n")
        output.write("\\.\n\n")

        output.write("COPY content_types (content_type) FROM stdin;\n")
        content_types.sort()
        for content_type in content_types:
            output.write(f"{content_type}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_content_types (m_id, content_type) FROM stdin;\n")
        for key in has_content_type.keys():
            output.write(f"{key}\t{has_content_type[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_details_locations (m_id, details, locations) FROM stdin;\n")
        for row in location_details:
            output.write(f"{row[0]}\t{row[1]}\t{row[2]}\n")
        output.write("\\.\n\n")

def handle_content_detail(location_details,m_id, content_detail, content_location):
    if not isinstance(content_location, str):
        content_location = f'{content_location}'
        if content_location.endswith('.0'):
            content_location = content_location[0:-2]
    md = pattern.findall(f'{content_detail}')
    md_2 = pattern.findall(content_location)
    if len(md) != len(md_2):
        stderr(f"{m_id}\t{content_detail}\t{content_location}")
        location_details.append([m_id,content_detail,content_location])
        return
    for li in md:
        for ti in li:
            pass
    for li in md_2:
        for ti in li:
            ti_2 = ti.strip(' +)(')
            if ti_2:
                pass
    for i in range(0,len(md)):
        if len(md[i]) != len(md_2[i]):
            pass
        else:
            for j in range(0,len(md[i])):
                content = md[i][j].split(r'+')
                location = md_2[i][j].split(r'+')
                if len(content) == len(location):
                    for k in range(0,len(content)):
                        if content[k].strip(' +)(') or location[k].strip(' +)('):
                            stderr(f"{m_id}\t{content[k].strip(' +)(')}\t{location[k].strip(' +)(')}")
                            location_details.append([m_id,content[k].strip(' +)('),location[k].strip(' +)(')])
                else:
                    if len(content) > 1:
                        for k in range(0,len(content)):
                            stderr(f"{m_id}\t{content[k].strip(' +)(')}\t{location[0].strip(' +)(')}")
                            location_details.append([m_id,content[k].strip(' +)('),location[0].strip(' +)(')])
                    pass
    if len(md) == 0:
        stderr(f'{m_id}\t{content_detail}\t{content_location}')
        location_details.append([m_id,content_detail,content_location])

def create_schema(headers):
    create_table("manuscripts", headers)
    #
    create_table("scaled_places",
            ["place text primary key"])
    #
    create_table("manuscripts_scaled_places",
            ["m_id text references manuscripts(ID)",
                "place text references scaled_places(place)"])
    #
    create_table("scaled_dates",
            ["date text primary key"])
    #
    create_table("manuscripts_scaled_dates",
            ["m_id text references manuscripts(ID)",
                "date text references scaled_dates(date)"])
    #
    create_table("books",
            ["id int primary key", "roman text"])
    #
    create_table("manuscripts_books_included",
            ["m_id text references manuscripts(ID)",
                "b_id int references books(id)"])
    #
    create_table("content_types",
            ["content_type text primary key"]) 
    #
    create_table("manuscripts_content_types",
            ["m_id text references manuscripts(ID)",
                "content_type text references content_types(content_type)"])
    #
    create_table("manuscripts_details_locations",
            ["m_id text references manuscripts(ID)",
                "details text",
                "locations text"])

def create_table(table, columns):
     schema_out.write(f"DROP TABLE {table} CASCADE;\n")
     schema_out.write(f"CREATE TABLE {table} (\n        ")
     schema_out.write(",\n        ".join(columns))
     schema_out.write("\n);\n\n")


def try_roman(text):
    try:
        n = roman.fromRoman(text)
        return [n]
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

def getBGColor(book, sheet, row, col):
    xfx = sheet.cell_xf_index(row, col)
    xf = book.xf_list[xfx]
    bgx = xf.background.pattern_colour_index
    pattern_colour = book.colour_map[bgx]
    return pattern_colour

def clean_string(value):
    return value


def stderr(text=""):
    sys.stderr.write("{}\n".format(text))

def arguments():
    ap = argparse.ArgumentParser(description='Read file (csv, xls(x), sql_dump to make xml-rdf')
    ap.add_argument('-i', '--inputfile',
                    help="inputfile",
                    default = "20200227_manuscripts_mastersheet_CURRENT.xlsx")
    ap.add_argument('-o', '--outputfile',
                    help="outputfile",
                    default = "isidore_data.sql")
    ap.add_argument('-s', '--schema',
                    help="schema file",
                    default = "isidore_schema.sql")
    ap.add_argument('-q', '--quotechar',
                    help="quotechar",
                    default = "'" )
    ap.add_argument('-t', '--headerrow',
                    help="headerrow; 0=row 1 (default = 0)",
                    default = 0)
    args = vars(ap.parse_args())
    return args


def end_prog(code=0):
    stderr("einde: {}".format(datetime.today().strftime("%H:%M:%S")))
    sys.exit(code)

 
if __name__ == "__main__":
    stderr("start: {}".format(datetime.today().strftime("%H:%M:%S")))

    args = arguments()
    inputfiles = args['inputfile'].split(',')
    outputfile = args['outputfile']
    schemafile = args['schema']
    headerrow = args['headerrow']
    quotechar = args['quotechar']

    output = open(outputfile, "w", encoding="utf-8")
    schema_out = open(schemafile, "w", encoding="utf-8")
 
    xls_file(inputfiles)
    

    end_prog(0)

