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
pattern_2 = re.compile(r'([^[]*)\[([^]]*)]([^[]*)')
patt = re.compile(r'([^+)(\][]+)')

linked_tables = ["place_scaled","date_scaled","books_included","content_type", "place_absolute","physical_state_scaled","script","designed_as"]

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
        scaled_places = {
			"Continent": 1,
			"England": 2,
			"France": 3,
			"northern France": 4,
			"southern France": 5,
			"German area": 6,
			"Ireland": 7,
			"Italy": 8,
			"central Italy": 9,
			"northern Italy": 10,
			"southern Italy": 11,
			"Spain": 12,
			"unknown": 13 }
        scaled_place_last_key = 13
        has_scaled_place = {}
        absolute_places = {}
        absolute_place_last_key = 0
        has_absolute_place = {}
        scaled_dates = {
		'7th c.': 70,
		'7th c., 2/2': 72,
		'8th c.': 80,
		'8th c., 1/2': 81,
		'8th c., 2/2': 82,
		'9th c.': 90,
		'9th c., 1/2': 91,
		'9th c., 2/2': 92,
		'10th c.': 100,
		'10th c., 1/2': 101,
		'10th c., 2/2': 102,
		'11th c.': 110,
		'11th c., 1/2': 111 }
        has_scaled_date = {}
        scaled_date_last_key = 119
        content_types = {}
        has_content_type = {}
        content_type_last_key = 0
        physical_states = []
        has_physical_state = {}
        scripts = {}
        has_script = {}
        scripts_last_key = 0
        designed_as = {}
        has_designed_as = {}
        designed_as_last_key = 0
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
                        # works only for .xls not for .xlsx
                        # stderr(f'{cell} ({rownum},{colnum}): {color}')
                        pass
                if cell.strip() != '':
                    if cell_type==xlrd.XL_CELL_TEXT:
                        cell = re.sub(r'&','&amp;',cell)
                        cell = re.sub(r"\\","\\\\\\\\",cell)
                        manuscript.append(cell.replace('\n',' '))
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
                        if cell=='Central Italy':
                            cell='central Italy'
                        if not cell in scaled_places:
                            scaled_place_last_key += 1
                            scaled_places[cell] = scaled_place_last_key
                        has_scaled_place[m_id] = scaled_places.get(cell)
                        manuscript.pop()
                    elif headers[colnum]=="place_absolute":
                        if not cell in absolute_places:
                            absolute_place_last_key += 1
                            absolute_places[cell] = absolute_place_last_key 
                        has_absolute_place[m_id] = absolute_places.get(cell)
                        manuscript.pop()
                    elif headers[colnum]=="date_scaled":
                        if not cell in scaled_dates:
                            scaled_date_last_key += 1
                            scaled_dates[cell] = scaled_date_last_key
                        has_scaled_date[m_id] = scaled_dates.get(cell)
                        manuscript.pop()
                    elif headers[colnum]=="books_included":
                        res = try_roman(cell)
                        includes_books[m_id] = res
                        manuscript.pop()
                    elif headers[colnum]=="content_type":
                        if not cell in content_types:
                            content_type_last_key += 1
                            content_types[cell] = content_type_last_key 
                        has_content_type[m_id] = content_types.get(cell)
                        manuscript.pop()
                    elif headers[colnum]=="physical_state_scaled":
                        if not cell in physical_states:
                            physical_states.append(cell)
                        has_physical_state[m_id] = cell
                        manuscript.pop()
                    elif headers[colnum]=="script":
                        if not cell in scripts:
                            scripts_last_key += 1
                            scripts[cell] = scripts_last_key
                        has_script[m_id] = scripts.get(cell)
                        manuscript.pop()
                    elif headers[colnum]=="designed_as":
                        # cell nog splitsen op komma
                        split_cell = cell.split('+')
                        for cel in split_cell:
                            if not cel.strip() in designed_as:
                                designed_as_last_key += 1
                                designed_as[cel.strip()] = designed_as_last_key
                            if m_id in has_designed_as:
                                has_designed_as[m_id].append(designed_as.get(cel.strip()))
                            else:
                                has_designed_as[m_id] = [designed_as.get(cel.strip())]
                        manuscript.pop()
                elif not headers[colnum] in linked_tables:
                    if colnum>12 and colnum<24:
                        manuscript.append('\\N')
                    else:
                        manuscript.append('')
            handle_content_detail(location_details, m_id, sheet.cell_value(rownum,27), sheet.cell_value(rownum,28))
            result.append(manuscript)
            teller += 1

        headers_2 = []
        for header in headers:
            if not header in linked_tables:
            # ["place_scaled","date_scaled","books_included","content_type"]:
                headers_2.append(header + " text")
        headers_2[0] += " primary key"
        create_schema(headers_2)

        headers_2 = []
        for header in headers:
            if not header in linked_tables:
            # ["place_scaled","date_scaled","books_included","content_type"]:
                headers_2.append(header)
        output.write("COPY manuscripts (")
        output.write(", ".join(headers_2))
        output.write(") FROM stdin;\n")
        for row in result:
#            row_str = "\t".join(row).replace('"','')
            output.write("\t".join(row).replace("'","''"))
            output.write("\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_places (place_id, place) FROM stdin;\n")
        for (place, place_id) in scaled_places.items():
            output.write(f"{place_id}\t{place}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scaled_places (m_id, place_id) FROM stdin;\n")
        for key in has_scaled_place.keys():
            output.write(f"{key}\t{has_scaled_place[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY absolute_places (place_id, place) FROM stdin;\n")
        for (place, place_id) in absolute_places.items():
            output.write(f"{place_id}\t{place}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_absolute_places (m_id, place_id) FROM stdin;\n")
        for key in has_absolute_place.keys():
            output.write(f"{key}\t{has_absolute_place[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_dates (date_id, date) FROM stdin;\n")
        for (date,date_id) in scaled_dates.items():
            output.write(f"{date_id}\t{date}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scaled_dates (m_id, date_id) FROM stdin;\n")
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

        output.write("COPY content_types (type_id, content_type) FROM stdin;\n")
        for (content_type, type_id) in content_types.items():
            output.write(f"{type_id}\t{content_type}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_content_types (m_id, type_id) FROM stdin;\n")
        for key in has_content_type.keys():
            output.write(f"{key}\t{has_content_type[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY scripts (script_id, script) FROM stdin;\n")
        for (script,script_id) in scripts.items():
            output.write(f"{script_id}\t{script}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scripts (m_id, script_id) FROM stdin;\n")
        for key in has_script.keys():
            output.write(f"{key}\t{has_script[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY designed_as (design_id, design) FROM stdin;\n")
        for (design,design_id) in designed_as.items():
            output.write(f"{design_id}\t{design}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_designed_as (m_id, design_id) FROM stdin;\n")
        for key in has_designed_as.keys():
            for design in has_designed_as[key]:
                output.write(f"{key}\t{design}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_details_locations (m_id, details, locations) FROM stdin;\n")
        for row in location_details:
            output.write(f"{row[0]}\t{row[1]}\t{row[2]}\n")
        output.write("\\.\n\n")


def string_to_dict(text):
    text_spl = str(text).split('+')
    for i in range(0,len(text_spl)):
        text_spl[i] = str(text_spl[i]).strip().replace(' ', '_')
    text = "+".join(text_spl)
    res = ''
    try:
        res = '[' + patt.sub(r'"\1"', text) + ']'
        res = res.replace(r"+", ",").replace('(','[').replace(')',']').replace("_",' ')
        return json.loads(res)
    except TypeError:
        return json.loads('["' + str(res) + '"]')
    except json.decoder.JSONDecodeError:
        return None

def handle_content_detail(location_details,m_id, content_details, content_locations):
    if not (content_details and content_locations):
        stderr(f"None: {m_id}\t{content_details}\t{content_locations}")
        if not content_details:
            content_details = "unknown"
        else:
            content_locations = "unknown"
    res_det = string_to_dict(content_details)
    res_loc = string_to_dict(content_locations)
    if not (res_det and res_loc):
        stderr(f"error: {m_id}")
        stderr(f"     : \t{content_details}")
        stderr(f"     : \t{content_locations}")
        add_location_details(location_details,m_id, content_details[1:-1], content_locations[1:-1])
    else:
        res = add_location_details(location_details,m_id, res_det, res_loc)
        if not res:
            add_location_details(location_details,m_id, content_details, content_locations)
    return

def add_location_details(location_details,m_id, res_det, res_loc):
    if isinstance(res_det, str) and isinstance(res_loc, str):
        location_details.append([m_id, res_det, res_loc])
        return True
    if isinstance(res_det, list) and isinstance(res_loc, list):
        if len(res_det)==1 and len(res_loc)>1:
            flat_det = flatten(res_det)
            flat_loc = flatten(res_loc)
            for loc in flat_loc:
                location_details.append([m_id, flat_det[0], loc])
        elif len(res_det)>1 and len(res_loc)==1:
            flat_det = flatten(res_det)
            flat_loc = flatten(res_loc)
            for det in flat_det:
                location_details.append([m_id, det, flat_loc[0]])
        elif len(res_det) != len(res_loc):
            return False
        elif len(res_det) == len(res_loc):
            for i in range(0, len(res_loc)):
                add_location_details(location_details, m_id, res_det[i], res_loc[i])
    elif isinstance(res_loc, str):
        for det in res_det:
            add_location_details(location_details, m_id, det, res_loc)
    elif isinstance(res_det, str):
        for loc in res_loc:
            add_location_details(location_details, m_id, res_det, loc)
    return True


def flatten(lijst):
    res = []
    for it in lijst:
        if isinstance(it, list):
            res += flatten(it)
        else:
            res.append(it)
    return res


def create_schema(headers):
    create_table("manuscripts", headers)
    #
    create_table("scaled_places",
            ["place_id integer primary key", "place text", "longitude real",
                "lattitude real"])
    #
    create_table("manuscripts_scaled_places",
            ["m_id text unique references manuscripts(ID)",
                "place_id integer references scaled_places(place_id)"])
    #
    create_table("absolute_places",
            ["place text primary key"])
    #
    create_table("manuscripts_absolute_places",
            ["m_id text unique references manuscripts(ID)",
                "place text references absolute_places(place),",
                "longitude real,",
                "lattitude real"])
    #
    create_table("scaled_dates",
            ["date text primary key"])
    #
    create_table("manuscripts_scaled_dates",
            ["m_id text unique references manuscripts(ID)",
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
            ["type_id integer primary key", 
                "content_type text"]) 
    #
    create_table("manuscripts_content_types",
            ["m_id text unique references manuscripts(ID)",
                "content_type integer references content_types(type_id)"])
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

def clean(value):
    return value.strip(' +)(][')


def stderr(text=""):
    sys.stderr.write("{}\n".format(text))

def arguments():
    ap = argparse.ArgumentParser(description='Read isidore xlsx to make postgres import file')
    ap.add_argument('-i', '--inputfile',
                    help="inputfile",
                    default = "20200615_manuscripts_mastersheet.xlsx")
    ap.add_argument('-o', '--outputfile',
                    help="outputfile",
                    default = "isidore_data_20200615.sql")
    ap.add_argument('-s', '--schema',
                    help="schema file",
                    default = "isidore_schema_2020615.sql")
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

