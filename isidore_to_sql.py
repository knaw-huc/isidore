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
long_lat_patt = re.compile(r"([NSEW]) (\d+)° (\d+)' (\d+)'?'?")

linked_tables = ["place_scaled","date_scaled","books_included","content_type", "place_absolute","physical_state_scaled","script","designed_as", "certainty"]

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
#
        scaled_places = get_scaled_places(wb)
        absolute_places = get_absolute_places(wb)
        viaf = get_viaf(wb)
        librarys, manuscripts_librarys = get_current_locations(wb)
#
        sheet = wb.sheet_by_index(0) 
        result = []
        headers = []
        scaled_place_last_key = 13
        has_scaled_place = {}
        absolute_place_last_key = 0
        has_absolute_place = {}
        scaled_dates = {
		'7th c.': [70,600,699],
		'7th c., 2/2': [72,650,699],
		'8th c.': [80,700,799],
		'8th c., 1/2': [81,700,749],
		'8th c., 2/2': [82,750,799],
		'9th c.': [90,800,899],
		'9th c., 1/2': [91,800,849],
		'9th c., 2/2': [92,850,899],
		'10th c.': [100,900,999],
		'10th c., 1/2': [101,900,949],
		'10th c., 2/2': [102,950,999],
		'11th c.': [110,1000,1099],
		'11th c., 1/2': [111,1000,1049] }
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
                            scaled_places[cell] = [scaled_place_last_key]
                        has_scaled_place[m_id] = scaled_places.get(cell)[0]
                        manuscript.pop()
                    elif headers[colnum]=="place_absolute":
                        if cell=='Chabannes/Limoges':
                            cell = 'Chabannes or Limoges'
                        if cell=='Rhaetia':
                            cell = 'Raetia'
                        if not cell in absolute_places:
                            absolute_place_last_key += 1
                            absolute_places[cell] = absolute_place_last_key
                        has_absolute_place[m_id] = [absolute_places.get(cell)[0]]
                        has_absolute_place[m_id].append(f"{sheet.cell_value(rownum,colnum+1)}")
                        manuscript.pop()
                    elif headers[colnum]=="certainty":
                        manuscript.pop()
                    elif headers[colnum]=="date_scaled":
                        if not cell in scaled_dates:
                            scaled_date_last_key += 1
                            scaled_dates[cell] = scaled_date_last_key
                        has_scaled_date[m_id] = scaled_dates.get(cell)[0]
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
            handle_content_detail(location_details, m_id, sheet.cell_value(rownum,33), sheet.cell_value(rownum,34))
            result.append(manuscript)
            teller += 1

        headers_2 = []
        for header in headers:
            if not header in linked_tables:
                headers_2.append(header + " text")
        headers_2[0] += " primary key"
        create_schema(headers_2)

        headers_2 = []
        for header in headers:
            if not header in linked_tables:
                headers_2.append(header)
        output.write("COPY manuscripts (")
        output.write(", ".join(headers_2))
        output.write(") FROM stdin;\n")
        for row in result:
            output.write("\t".join(row)) # .replace("'","''"))
            output.write("\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_places (place_id, place, gps_latitude, gps_longitude, latitude, longitude) FROM stdin;\n")
        for (place, place_data) in scaled_places.items():
            output.write(f"{place_data[0]}\t{place}\t{place_data[1]}\t{place_data[2]}\t{place_data[5]}\t{place_data[6]}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_scaled_places (m_id, place_id) FROM stdin;\n")
        for key in has_scaled_place.keys():
            output.write(f"{key}\t{has_scaled_place[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY absolute_places (place_id, place_absolute, GPS_latitude, GPS_longitude, Country, Country_GeoNames, Latitude, Longitude, GeoNames_id, GeoNames_uri) FROM stdin;\n")
        for (place, place_data) in absolute_places.items():
            all_place_data = '\t'.join(place_data[2:])
            output.write(f"{place_data[0]}\t{all_place_data}\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_absolute_places (m_id, place_id, certainty) FROM stdin;\n")
        for key in has_absolute_place.keys():
            output.write(f"{key}\t{has_absolute_place[key][0]}\t{has_absolute_place[key][1]}\n")
        output.write("\\.\n\n")

        output.write("COPY scaled_dates (date_id, date, lower_date, upper_date) FROM stdin;\n")
        for (date,date_id) in scaled_dates.items():
            output.write(f"{date_id[0]}\t{date}\t{date_id[1]}\t{date_id[2]}\n")
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

        output.write("COPY designed_as (design_id, designed_as) FROM stdin;\n")
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

        output.write("COPY manuscripts_viaf (ID, shelfmark, additional_content_scaled, VIAF_ID, VIAF_URL, Full_name_1, Full_name_2) FROM stdin;\n")
        for row in viaf:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY library (lib_id, lib_name, GPS_latitude, GPS_longitude, Place_name, Country, Country_GeoNames, Latitude, Longitude, GeoNames_id, GeoNames_uri) FROM stdin;\n")
        for key in librarys:
            output.write("\t".join(librarys[key]) + "\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_library (m_id, shelfmark, lib_id) FROM stdin;\n")
        for row in manuscripts_librarys:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")


def get_scaled_places(wb):
    teller = 0
    scaled_places = {}
    sheet = wb.sheet_by_index(5) 
    for rownum in range(1, sheet.nrows):
        placename = ''
        for colnum in range(1, sheet.ncols):
            cell_type = sheet.cell_type(rownum,colnum)
            cell = "{}".format(sheet.cell_value(rownum,colnum))
            if cell=='Central Italy':
#           Central = central
                cell = 'central Italy'
            if colnum==1:
                if not cell in scaled_places:
                    placename = cell
                    scaled_places[placename] = [0]
            elif placename != '':
                scaled_places[placename].append(cell.replace('"',"''"))
        if placename and scaled_places[placename][1]:
            scaled_places[placename].append(hms_to_dec(scaled_places[placename][1]))
            scaled_places[placename].append(hms_to_dec(scaled_places[placename][2]))
        elif placename:
            scaled_places[placename].append('0.0')
            scaled_places[placename].append('0.0')
#           sorteer op alfabet
    for k in sorted(scaled_places.keys()):
        teller += 1
        scaled_places[k][0] = teller
    return scaled_places


def get_absolute_places(wb):
    teller = 0
    absolute_places = {}
    sheet = wb.sheet_by_index(4) 
    for rownum in range(1, sheet.nrows):
        placename = ''
        cell = "{}".format(sheet.cell_value(rownum,2))
        if not cell in absolute_places:
            placename = cell
            absolute_places[placename] = [0]
        for colnum in range(1, sheet.ncols):
            cell_type = sheet.cell_type(rownum,colnum)
            cell = "{}".format(sheet.cell_value(rownum,colnum))
            if placename != '':
                if colnum==7 or colnum==8:
                    if cell=='':
                        absolute_places[placename].append('0.0')
                    else:
                        absolute_places[placename].append(cell.replace('"',"''"))
                else:
                    absolute_places[placename].append(cell.replace('"',"''"))
    for k in sorted(absolute_places.keys()):
        teller += 1
        absolute_places[k][0] = teller
    return absolute_places


def get_viaf(wb):
    viaf = []
    sheet = wb.sheet_by_index(2) 
    for rownum in range(1, sheet.nrows):
        placename = ''
        row = []
        for colnum in range(sheet.ncols):
            cell = "{}".format(sheet.cell_value(rownum,colnum))
            if colnum==3:
                if cell.endswith('.0'):
                    cell = cell[0:-2]
                elif cell=='':
                    cell = '0'
#                cell = f'"{cell}"'
            row.append(cell)
        viaf.append(row)
    return viaf


def get_current_locations(wb):
    librarys = {}
    teller = 0
    manuscripts_librarys = []
    sheet = wb.sheet_by_index(3)
    for rownum in range(1, sheet.nrows):
        cell = "{}".format(sheet.cell_value(rownum,1))
        library = ','.join(cell.split(',')[0:2])
        if library=='Ithaca, Cornell University' or library=='Salzburg, St. Peter':
            library = ','.join(cell.split(',')[0:3])
        if library not in librarys:
            teller += 1
            librarys[library] = [str(teller), library]
            for colnum in range(2, sheet.ncols):
                cell = "{}".format(sheet.cell_value(rownum,colnum))
                if colnum==9:
                    if cell.endswith('.0'):
                        cell = cell[0:-2]
                    if cell=='':
                        cell = '0'
                librarys[library].append(cell)
        manuscripts_librarys.append( [ "{}".format(sheet.cell_value(rownum,0)),
            "{}".format(sheet.cell_value(rownum,1)), librarys[library][0] ] )
    return librarys, manuscripts_librarys


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
        if not content_details:
            content_details = "unknown"
        if not content_locations:
            content_locations = "unknown"
#        stderr(f"{m_id}\tcontent_details:\t{content_details}\n\tcontent_locations:\t{content_locations}")
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
            stderr(f'{m_id}: possible mismatch between details and location:\
                    \n{content_details}\n{content_locations}')
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
#            stderr(f'{m_id}: possible mismatch between details and location:\n{flat_det}\n{flat_loc}')
            for loc in flat_loc:
                location_details.append([m_id, flat_det[0], loc])
        elif len(res_det)>1 and len(res_loc)==1:
            flat_det = flatten(res_det)
            flat_loc = flatten(res_loc)
#            stderr(f'{m_id}: possible mismatch between details and location:\n{flat_det}\n{flat_loc}')
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
            ["place_id integer primary key", "place text",
                "gps_latitude text", "gps_longitude text",
                "latitude real", "longitude real"])
    #
    create_table("manuscripts_scaled_places",
            ["m_id text unique references manuscripts(ID)",
                "place_id integer references scaled_places(place_id)"])
    #
    create_table("absolute_places",
            ["place_id integer primary key", "place_absolute text",
                "GPS_latitude text", "GPS_longitude text", "Country text", "Country_GeoNames text",
                "Latitude real", "Longitude real", "GeoNames_id text", "GeoNames_uri text"])
    #
    create_table("manuscripts_absolute_places",
            ["m_id text unique references manuscripts(ID)",
                "place_id integer references absolute_places(place_id)",
                "certainty text"])
    #
    create_table("scaled_dates",
            ["date_id integer primary key","date text", "lower_date integer","upper_date integer"])
    #
    create_table("manuscripts_scaled_dates",
            ["m_id text unique references manuscripts(ID)",
                "date_id integer references scaled_dates(date_id)"])
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
                "type_id integer references content_types(type_id)"])
    #
    create_table("scripts",
            ["script_id integer primary key", "script text"])
    #
    create_table("manuscripts_scripts",
            ["m_id text references manuscripts(ID)", "script_id integer references scripts(script_id)"])
    #
    create_table("designed_as",
            ["design_id integer primary key", "designed_as text"])
    #
    create_table("manuscripts_designed_as",
            ["m_id text references manuscripts(ID)", "design_id integer references designed_as(design_id)"])
    #
    create_table("manuscripts_details_locations",
            ["m_id text references manuscripts(ID)",
                "details text",
                "locations text"])
    #
    schema_out.write('/* VIAF_ID is text: some of these ids have more digits than either integer or bigint can handle! */\n')
    create_table("manuscripts_viaf",
            ["ID text references manuscripts(ID)", "shelfmark text",
                "additional_content_scaled text", "VIAF_ID text",
                "VIAF_URL text", "Full_name_1 text", "Full_name_2 text"])
    #
    create_table("library",
            ["lib_id integer primary key", "lib_name text",
                "GPS_latitude text", "GPS_longitude text",
                "Place_name text", "Country text", "Country_GeoNames text",
                "Latitude real", "Longitude real", "GeoNames_id integer", "GeoNames_uri text"])
    #
    create_table("manuscripts_library",
            ["m_id text references manuscripts(ID)", "shelfmark text", "lib_id integer references library(lib_id)"])


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

# convert longitude or latitude in degrees, minutes, seconds to degrees and a decimal fraction
def hms_to_dec(text):
    res = re.search(long_lat_patt, text)
    res_float = float(res.group(2)) + float(res.group(3)) / 60 + float(res.group(4)) / 3600
    # South and West are negative values:
    if res.group(1)=='S' or res.group(1)=='W':
        res_float = -1 * res_float 
    # a decimal fraction of maximum 5 digits seems to be the international standard
    return f'{round(res_float,5):.5f}'

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

