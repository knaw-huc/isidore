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

# This script uses xlrd to read xlsx files.
# According to https://xlrd.readthedocs.io/en/latest/index.html :
# 'This library will no longer read anything other than .xls files'
# In future a conversion to another library might be nescessary, for example: openpyxl
# (https://foss.heptapod.net/openpyxl)

output = None
schema_out = None
delimiter = ','
quotechar = ''
pattern = re.compile(r'([^(]*)\(([^)]*)\)([^(]*)')
pattern_2 = re.compile(r'([^[]*)\[([^]]*)]([^[]*)')
patt = re.compile(r'([^+)(\][]+)')
long_lat_patt = re.compile(r"([NSEW]) (\d+)° (\d+)' (\d+)'?'?")
manuscript_ids = []

# linked_tables: a list of columns in Mastersheet that are not added to the manuscripts table
# in the database. There are other tabs with the same information as in these columns, all
# linking to the manuscripts table.
linked_tables = ["place_scaled", "date_scaled", "books_included", "content_type",
        "place_absolute",
        #"physical_state_scaled",
        "script", "designed_as", "certainty",
        "source_of_dating", "provenance_scaled", "related_mss_in_the_database",
        "related_mss_outside_of_the_database", "reason_for_relationship", "annotations", "diagrams"]

# ignore: a list of columns in mastersheet to be ignored. The data in these columns also
# exists in other tabs.
ignore = ["certainty", "related_mss_in_the_database", "related_mss_outside_of_the_database",
        "reason_for_relationship", "annotations", "diagrams"]
 

def xls_file(inputfiles, headerrow=0):
    other_headers = {}
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
        get_manuscript_ids(wb)
        scaled_places = get_scaled_places(wb)
        absolute_places,absolute_place_last_key = get_absolute_places(wb)
        source_of_dating = get_source_of_dating(wb)
        provenance_scaled = get_provenance_scaled(wb)
        location_details = get_location_details(wb)
        viaf,other_headers['viaf'] = get_viaf(wb)
        librarys, manuscripts_librarys = get_current_locations(wb)
        relationships = get_relationships(wb)
        interpolations = get_interpolations(wb)
        stderr(f'lengte interpolations: {len(interpolations[0])}')
        diagrams = get_diagrams(wb)
        easter_table = get_easter_table(wb)
        annotations = get_annotations(wb)
        urls = get_urls(wb)
#
        logfile = open('cdl_1.log','w')
        for row in location_details:
            logfile.write(f'{row}\n')
        logfile.close()
#
        result = []
        headers = []
        # next line: change this to something durable
        scaled_place_last_key = 13
        has_scaled_place = {}
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
        has_source_of_dating = {}
        has_provenance_scaled = {}
        designed_as_last_key = 0
        includes_books = {}

        sheet = wb.sheet_by_name('Mastersheet') 
        headers = get_column_headers(sheet,headerrow)
#        for colnum in range(sheet.ncols):
#            if sheet.cell_value(headerrow,colnum) != '':
#                headers.append(re.sub(r'[ -/]+','_',sheet.cell_value(headerrow,colnum)).strip('_'))
#            else:
#                headers.append(f"empty{colnum}")
        teller = 0
        for rownum in range((headerrow+1), sheet.nrows):
            manuscript = []
            m_id = f"{sheet.cell_value(rownum,0)}"
            for colnum in range(sheet.ncols):
                cell_type = sheet.cell_type(rownum,colnum)
                cell = f"{sheet.cell_value(rownum,colnum)}"
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
                    elif cell_type==xlrd.XL_CELL_DATE:
                        if cell.endswith('.0'):
                            cell = cell[0:-2]
                        date = xlrd.xldate.xldate_as_datetime(int(cell),0)
                        try:
                            manuscript.append(f'{date.strftime("%Y-%m-%d")}')
                        except:
                            stderr(f'cell: {cell} (colnum: {colnum})')
                            end_prog(1)
                    else:
                        stderr(f'Cell type: {cell_type} (Colnum: {colnum}) Not found')
                    if headers[colnum]=="place_scaled":
                        if cell=='Central Italy':
                            cell='central Italy'
                        if not cell in scaled_places:
                            scaled_place_last_key += 1
                            scaled_places[cell] = [scaled_place_last_key]
                        has_scaled_place[m_id] = scaled_places.get(cell)[0]
                        manuscript.pop()
                    elif headers[colnum]=="place_absolute":
                        if cell=='Limoges/Angoulême':
                            cell = 'Limoges or Angoulême'
                        if cell=='Chabannes/Limoges':
                            pass
#                            cell = 'Chabannes or Limoges'
                        if cell=='Rhaetia':
                            pass
#                            cell = 'Raetia'
                        if not cell in absolute_places:
                            stderr(f'add to absolute_places: {cell}')
                            absolute_place_last_key += 1
                            stderr(absolute_places)
                            absolute_places[cell] = [absolute_place_last_key,cell,'','','','','','0.0','0.0','','']
                            stderr(absolute_places)
                        try:
                            has_absolute_place[m_id] = [absolute_places.get(cell)[0]]
                            has_absolute_place[m_id].append(f"{sheet.cell_value(rownum,colnum+1)}")
                            manuscript.pop()
                        except:
                            stderr(f'cell: {cell}')
                            stderr(f'abs?: {absolute_places[cell]}')
                            end_prog(1)
                    elif headers[colnum] in ignore:
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
                        #manuscript.pop()
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
                    elif headers[colnum]=="source_of_dating":
                        has_source_of_dating[m_id] = source_of_dating.index(cell.strip())
                        manuscript.pop()
                    elif headers[colnum]=="provenance_scaled":
                        for prov in re.split(';',cell):
                            if m_id in has_provenance_scaled:
                                has_provenance_scaled[m_id].append(provenance_scaled.index(prov.strip()))
                            else:
                                has_provenance_scaled[m_id] = [ provenance_scaled.index(prov.strip()) ]
                        manuscript.pop()
                elif not headers[colnum] in linked_tables:
                    # pas dit aan (maar wat doet dit eigenlijk?)
                    # er was een reden om bepaalde lege velden te vullen
                    # met een \N ipv '', maar nu (18-02-21) vergeten waarom.
                    if colnum>12 and colnum<24:
                        manuscript.append('\\N')
                    else:
                        manuscript.append('')
            result.append(manuscript)
            teller += 1

        headers_2 = []
        for header in headers:
            if not header in linked_tables:
                headers_2.append(header + " text")
        headers_2[0] += " primary key"
        create_schema(headers_2,other_headers)

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

    # think of something to get the headers from the xlsx
    # instead of adjusting them manually when (once again)
    # the sheet is changed
    # see also: create_schema
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

        logfile = open('cdl_2.log','w')
        for row in location_details:
            logfile.write(f'{row}\n')
        output.write("COPY manuscripts_details_locations (m_id, material_type, books_included, details, locations) FROM stdin;\n")
        for row in location_details:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        # viaf
        header_string = ", ".join(other_headers['viaf'])
        output.write(f"COPY manuscripts_viaf ({header_string}) FROM stdin;\n")
#        output.write(f"COPY manuscripts_viaf (ID, shelfmark, additional_content_scaled, VIAF_ID, VIAF_URL, Full_name_1, Full_name_2,Biblissima_author_URL, Wikidata_author_url) FROM stdin;\n")
        # stderr('Attention:')
        # stderr('due to 3 empty last columns (table VIAF) row has to be shortened!!!')
        # repaired in xlsx (3-12-2021)
        for row in viaf:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY library (lib_id, lib_name, GPS_latitude, GPS_longitude, Place_name, Country, Country_GeoNames, Latitude, Longitude, GeoNames_id, GeoNames_uri) FROM stdin;\n")
        for key in librarys:
            output.write("\t".join(librarys[key]) + "\n")
        output.write("\\.\n\n")

        output.write("COPY manuscript_current_places (m_id, shelfmark, lib_id) FROM stdin;\n")
        for row in manuscripts_librarys:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY source_of_dating (s_id, source) FROM stdin;\n")
        for tel in range(0,len(source_of_dating)):
            output.write(f'{tel}'+ "\t" + source_of_dating[tel] + "\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_source_of_dating (m_id, s_id) FROM stdin;\n")
        for key in has_source_of_dating:
            output.write(f"{key}\t{has_source_of_dating[key]}\n")
        output.write("\\.\n\n")

        output.write("COPY provenance_scaled (p_id, provenance) FROM stdin;\n")
        for tel in range(0,len(provenance_scaled)):
            output.write(f'{tel}'+ "\t" + provenance_scaled[tel] + "\n")
        output.write("\\.\n\n")

        output.write("COPY manuscripts_provenance_scaled (m_id, p_id) FROM stdin;\n")
        for key in has_provenance_scaled:
            for prov in has_provenance_scaled[key]:
                output.write(f"{key}\t{prov}\n")
        output.write("\\.\n\n")

        output.write("COPY relationships (m_id, shelfmark, rel_mss_id, rel_mss_other, reason) FROM stdin;\n")
        for row in relationships:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY interpolations (m_id, shelfmark, interpolation, folia, url, description) FROM stdin;\n")
        for row in interpolations:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY diagrams (m_id, shelfmark, diagram_type, folia, url, description) FROM stdin;\n")
        for row in diagrams:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY easter_table (m_id, shelfmark, easter_table_type, folia, remarks) FROM stdin;\n")
        for row in easter_table:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY annotations (m_id, shelfmark, number_of_annotations, amount, books, language, url, remarks) FROM stdin;\n")
        for row in annotations:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")

        output.write("COPY url (m_id, shelfmark, url_images, label, Biblissima_author_URL, mirabileweb, trismegistos, fama, manuscripta_medica, jordanus, bstk_online, handschriftencensus, dhbm, Bibliotheca_legum, Capitularia, other_links, label_other_links, iiif_manifest, copyright_status) FROM stdin;\n")
        for row in urls:
            output.write("\t".join(row) + "\n")
        output.write("\\.\n\n")


def get_column_headers(sheet,headerrow):
    headers = []
    for colnum in range(sheet.ncols):
        if sheet.cell_value(headerrow,colnum) != '':
            headers.append(re.sub(r'[ -/]+','_',sheet.cell_value(headerrow,colnum)).strip('_'))
        else:
            headers.append(f"empty{colnum}")
    return headers



def get_manuscript_ids(wb):
    sheet = wb.sheet_by_name('Mastersheet') 
    for rownum in range((headerrow+1), sheet.nrows):
        m_id = f"{sheet.cell_value(rownum,0)}"
        manuscript_ids.append(m_id)

def get_scaled_places(wb):
    # TODO: use the 0-column with the m-ids
    # but check the m-id in the table manuscript-ids
    # check_m_id(m_id, "Geo_placescaled", rownum):
    teller = 0
    scaled_places = {}
    sheet = wb.sheet_by_name('Geo_placescaled') 
    col_pl_name = sheet.row_values(0).index('place, scaled')
    for rownum in range(1, sheet.nrows):
        placename = ''
        for colnum in range(1, sheet.ncols):
            cell_type = sheet.cell_type(rownum,colnum)
            cell = f"{sheet.cell_value(rownum,colnum)}"
            if cell=='Central Italy':
#           Central = central
                cell = 'central Italy'
            if colnum==col_pl_name:
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
    for k in sorted(scaled_places.keys(), key=str.casefold):
        teller += 1
        scaled_places[k][0] = teller
    return scaled_places


def get_absolute_places(wb):
    # TODO: use the 0-column with the m-ids
    # but check the m-id in the table manuscript-ids
    # check_m_id(m_id, "Geo_placeabsolute", rownum):
    teller = 0
    absolute_places = {}
    sheet = wb.sheet_by_name('Geo_placeabsolute') 
    col_pl_name = sheet.row_values(0).index('place, absolute')
    col_lat = sheet.row_values(0).index('Latitude')
    col_long = sheet.row_values(0).index('Longitude')
    col_geo = sheet.row_values(0).index('GeoNames_id')
    for rownum in range(1, sheet.nrows):
        placename = ''
        cell = f"{sheet.cell_value(rownum,col_pl_name).strip()}"
        if not cell in absolute_places:
            placename = cell
            absolute_places[placename] = [0]
        for colnum in range(1, sheet.ncols):
            cell = f"{sheet.cell_value(rownum,colnum)}"
            if placename != '':
                if colnum==col_lat or colnum==col_long:
                    if cell=='':
                        absolute_places[placename].append('0.0')
                    else:
                        absolute_places[placename].append(cell.replace('"',"''"))
                elif colnum==col_geo:
                    if cell.endswith('.0'):
                        cell = cell[0:-2]
                    absolute_places[placename].append(cell)
                else:
                    cell = cell.strip()
                    absolute_places[placename].append(cell.replace('"',"''"))
    for k in sorted(absolute_places.keys(), key=str.casefold):
        teller += 1
        absolute_places[k][0] = teller
    return absolute_places,teller

def get_source_of_dating(wb):
    source_of_dating = []
    sheet = wb.sheet_by_name('Mastersheet')
    colnum = sheet.row_values(0).index('source of dating')
    for rownum in range(1, sheet.nrows):
        cell = f"{sheet.cell_value(rownum,colnum).strip()}"
        if not cell in source_of_dating and not cell=="":
            source_of_dating.append(cell)
    return sorted(source_of_dating, key=str.casefold)

def get_provenance_scaled(wb):
    provenance_scaled = []
    sheet = wb.sheet_by_name('Mastersheet')
    colnum = sheet.row_values(0).index('provenance, scaled')
    for rownum in range(1, sheet.nrows):
        cell = f"{sheet.cell_value(rownum,colnum).strip()}"
        cell_parts = re.split(';',cell)
        for cell_part in cell_parts:
            if not cell_part.strip() in provenance_scaled and not cell_part.strip()=="":
                provenance_scaled.append(cell_part.strip())
    return sorted(provenance_scaled, key=str.casefold)

def get_location_details(wb):
    location_details = []
    sheet = wb.sheet_by_name('content')
    col_loc = sheet.row_values(0).index('content location')
    col_det = sheet.row_values(0).index('content - detail')
    col_m_t = sheet.row_values(0).index('material type')
    col_b_i = sheet.row_values(0).index('books included')
    if(col_det<0 or col_loc<0):
        stderr('No content detail or content location available in sheet content')
        return []
    for rownum in range(1, sheet.nrows):
        m_id = f"{sheet.cell_value(rownum,0).strip()}"
        if not m_id in manuscript_ids:
            stderr(f'{m_id} in tab "content", row {rownum} not found on "Mastersheet": entry skipped')
            continue
        mat_type = f"{sheet.cell_value(rownum,col_m_t).strip()}"
        books_incl = f"{sheet.cell_value(rownum,col_b_i).strip()}"
        cont_detail = sheet.cell_value(rownum,col_det)
        if(sheet.cell_type(rownum,col_det)==xlrd.XL_CELL_NUMBER):
            cont_detail = str(cont_detail)
        else:
            cont_detail = cont_detail.strip()
        cont_location = sheet.cell_value(rownum,col_loc)
        if(sheet.cell_type(rownum,col_loc)==xlrd.XL_CELL_NUMBER):
            cont_location = str(cont_location)
            if(cont_location.endswith('.0')):
                cont_location = cont_location[0:-2]
        else:
            cont_location = cont_location.strip()
        handle_content_detail(location_details, m_id, mat_type, books_incl, cont_detail, cont_location)
    return location_details



def get_viaf(wb):
    viaf = []
    sheet = wb.sheet_by_name('VIAF') 
    headers = get_column_headers(sheet,0)
    col_viaf_id = sheet.row_values(0).index('VIAF ID')
    col_m_id = sheet.row_values(0).index('ID')
    for rownum in range(1, sheet.nrows):
        m_id = sheet.cell_value(rownum,col_m_id)
        if not check_m_id(m_id, "VIAF", rownum):
            continue
        placename = ''
        row = []
        for colnum in range(sheet.ncols):
            cell = f"{sheet.cell_value(rownum,colnum)}"
            if colnum==col_viaf_id:
                if cell.endswith('.0'):
                    cell = cell[0:-2]
                elif cell=='':
                    cell = '0'
            row.append(cell)
        viaf.append(row)
    return viaf,headers


def get_current_locations(wb):
    librarys = {}
    teller = 0
    manuscripts_librarys = []
    sheet = wb.sheet_by_name('Geo_currentlocation')
    col_shelf = sheet.row_values(0).index('Shelfmark')
    col_geo = sheet.row_values(0).index('GeoNames_id')
    col_m_id = sheet.row_values(0).index('ID')
    for rownum in range(1, sheet.nrows):
        m_id = sheet.cell_value(rownum,col_m_id)
        cell = f"{sheet.cell_value(rownum,col_shelf)}"
        library = ','.join(cell.split(',')[0:2])
        if library=='Ithaca, Cornell University' or library=='Salzburg, St. Peter':
            library = ','.join(cell.split(',')[0:3])
        if library not in librarys:
            teller += 1
            librarys[library] = [str(teller), library]
            if not check_m_id(m_id, "Geo_currentlocation", rownum):
                continue
            for colnum in range(2, sheet.ncols):
                cell = f"{sheet.cell_value(rownum,colnum)}"
                cell_type = sheet.cell_type(rownum,colnum)
                if not cell_type==xlrd.XL_CELL_NUMBER:
                    cell = cell.strip()
                if colnum==col_geo:
                    if cell.endswith('.0'):
                        cell = cell[0:-2]
                    if cell=='':
                        cell = '0'
                librarys[library].append(cell)
        manuscripts_librarys.append( [ f"{sheet.cell_value(rownum,col_m_id)}",
            f"{sheet.cell_value(rownum,col_shelf)}", librarys[library][0] ] )
    return librarys, manuscripts_librarys


def get_relationships(wb):
    pattern = re.compile(r'M\d\d\d\d')
    relationships = []
    sheet = wb.sheet_by_name('Relationships')
    rel_mss_colnum = sheet.row_values(0).index('related mss.')
    col_m_id = sheet.row_values(0).index('ID')
    for rownum in range(1, sheet.nrows):
        m_id = sheet.cell_value(rownum,col_m_id)
        if not check_m_id(m_id, "Relationships", rownum):
            continue
        relation = []
        for colnum in range(0,sheet.ncols-1):
            cell = sheet.cell_value(rownum,colnum).strip()
            if colnum==rel_mss_colnum:
                if pattern.match(cell):
                    if not check_m_id(cell, "Relationships", rownum):
                        continue
                    relation.append(cell)
                    relation.append("")
                else:
                    relation.append('\\N')
                    relation.append(cell)
            else:
                relation.append(cell)
        relationships.append(relation)
    return relationships


def get_interpolations(wb):
    return get_default(wb.sheet_by_name('Interpolations'))


def get_diagrams(wb):
    return get_default(wb.sheet_by_name('Diagrams'))


def get_easter_table(wb):
    return get_default(wb.sheet_by_name('EasterTable'))


def get_annotations(wb):
    return get_default(wb.sheet_by_name('Annotations'))


def get_urls(wb):
    return get_default(wb.sheet_by_name('URL'))


def get_default(sheet):
    defaults = []
    col_m_id = sheet.row_values(0).index('ID')
    for rownum in range(1, sheet.nrows):
        m_id = sheet.cell_value(rownum,col_m_id)
        if not check_m_id(m_id, sheet.name, rownum):
            continue
        relation = []
        for colnum in range(0,sheet.ncols):
            cell = sheet.cell_value(rownum,colnum)
            if sheet.cell_type(rownum,colnum)==xlrd.XL_CELL_TEXT:
                cell = cell.strip()
            elif sheet.cell_type(rownum,colnum)==xlrd.XL_CELL_NUMBER:
                if str(cell).endswith('.0'):
                    cell = str(cell)[0:-2]
            relation.append(cell)
        defaults.append(relation)
        if not len(relation)==sheet.ncols:
            stderr('lengte fout')
            stderr(relation)
            sys.exit(1)
    return defaults


def check_m_id(m_id,sheet_name,rownum):
    if not m_id in manuscript_ids:
        stderr(f'{m_id} on tab "{sheet_name}" row {rownum+1}, not found on "Mastersheet": entry skipped')
        return False
    return True


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

def handle_content_detail(location_details,m_id, mat_type, books_incl, content_details, content_locations):
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
        add_location_details(location_details,m_id, mat_type, books_incl, content_details[1:-1], content_locations[1:-1])
    else:
        res = add_location_details(location_details,m_id, mat_type, books_incl, res_det, res_loc)
        if not res:
            stderr(f'{m_id}: possible mismatch between details and location:\
                    \n{content_details}\n{content_locations}')
            add_location_details(location_details,m_id, mat_type, books_incl, content_details, content_locations)
    return

def add_location_details(location_details,m_id, mat_type, books_incl, res_det, res_loc):
    if isinstance(res_det, str) and isinstance(res_loc, str):
        location_details.append([m_id, mat_type, books_incl, res_det, res_loc])
        return True
    if isinstance(res_det, list) and isinstance(res_loc, list):
        if len(res_det)==1 and len(res_loc)>1:
            flat_det = flatten(res_det)
            flat_loc = flatten(res_loc)
#            stderr(f'{m_id}: possible mismatch between details and location:\n{flat_det}\n{flat_loc}')
            for loc in flat_loc:
                location_details.append([m_id, mat_type, books_incl, flat_det[0], loc])
        elif len(res_det)>1 and len(res_loc)==1:
            flat_det = flatten(res_det)
            flat_loc = flatten(res_loc)
#            stderr(f'{m_id}: possible mismatch between details and location:\n{flat_det}\n{flat_loc}')
            for det in flat_det:
                location_details.append([m_id, mat_type, books_incl, det, flat_loc[0]])
        elif len(res_det) != len(res_loc):
            return False
        elif len(res_det) == len(res_loc):
            for i in range(0, len(res_loc)):
                add_location_details(location_details, m_id, mat_type, books_incl, res_det[i], res_loc[i])
    elif isinstance(res_loc, str):
        for det in res_det:
            add_location_details(location_details, m_id, mat_type, books_incl, det, res_loc)
    elif isinstance(res_det, str):
        for loc in res_loc:
            add_location_details(location_details, m_id, mat_type, books_incl, res_det, loc)
    return True


def flatten(lijst):
    res = []
    for it in lijst:
        if isinstance(it, list):
            res += flatten(it)
        else:
            res.append(it)
    return res


def create_schema(headers,other_headers):
    # think of something to get the headers from the xlsx
    # instead of adjusting them manually when (once again)
    # the sheet is changed
    # see also:  line 265
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
            ["id int primary key", "roman text", "sections integer"])
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
            ["m_id text references manuscripts(ID)", "material_type text",
                "books_included text", "details text", "locations text"])
    #
    schema_out.write('/* VIAF_ID is text: some of these ids have more digits than either integer or bigint can handle! */\n')
    headers = list(map(lambda x: x + " text", other_headers['viaf'][1:]))
    create_table("manuscripts_viaf",
            ["ID text references manuscripts(ID)"] + headers)

#            , "shelfmark text",
#                "additional_content_scaled text", "VIAF_ID text",
#                "VIAF_URL text", "Full_name_1 text", "Full_name_2 text",
#                "Biblissima_author_URL text, Wikidata_author_url text"
#                ])
    #
    create_table("library",
            ["lib_id integer primary key", "lib_name text",
                "GPS_latitude text", "GPS_longitude text",
                "Place_name text", "Country text", "Country_GeoNames text",
                "Latitude real", "Longitude real", "GeoNames_id integer", "GeoNames_uri text"])
    #
    create_table("manuscript_current_places",
            ["m_id text references manuscripts(ID)", "shelfmark text", "lib_id integer references library(lib_id)"])
    #
    create_table("source_of_dating",
            ["s_id integer primary key","source text"])
    #
    create_table("manuscripts_source_of_dating",
            ["m_id text references manuscripts(ID)","s_id integer references source_of_dating(s_id)"])
    #
    create_table("provenance_scaled",
            ["p_id integer primary key","provenance text"])
    #
    create_table("manuscripts_provenance_scaled",
            ["m_id text references manuscripts(ID)","p_id integer references provenance_scaled(p_id)"])
    #
    create_table("relationships",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "rel_mss_id text references manuscripts(ID)", "rel_mss_other text",
                "reason text"])
    #
    create_table("interpolations",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "interpolation text", "folia text", "url text", "description text"])
    #
    create_table("diagrams",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "diagram_type text", "folia text", "url text", "description text"])
    #
    create_table("easter_table",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "easter_table_type text", "folia text", "remarks text"])
    #
    create_table("annotations",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "number_of_annotations text", "amount text", "books text",
                "language text", "url text", "remarks text"])
    #
    create_table("url",
            ["m_id text references manuscripts(ID)", "shelfmark text",
                "url_images text", "label text", "Biblissima_author_URL text", "mirabileweb text",
                "trismegistos text", "fama text", "manuscripta_medica text",
                "jordanus text", "bstk_online text", "handschriftencensus text",
                "dhbm text", "Bibliotheca_legum text", "Capitularia text", 
                "other_links text", "label_other_links text",
                "iiif_manifest text", "copyright_status text",
                "permission_from_the_library text", "image_published text"])
            
                

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


def stderr(text="",nl="\n"):
    sys.stderr.write(f"{text}{nl}")

def arguments():
    today = f'{datetime.today().strftime("%Y%m%d")}'
    ap = argparse.ArgumentParser(description='Read isidore xlsx to make postgres import file')
    ap.add_argument('-i', '--inputfile',
                    help="inputfile",
                    default = "manuscripts_mastersheet.xls" )
    ap.add_argument('-o', '--outputfile',
                    help="outputfile",
                    default = f"isidore_data_{today}.sql" )
    ap.add_argument('-s', '--schema',
                    help="schema file",
                    default = f"isidore_schema_{today}.sql" )
    ap.add_argument('-q', '--quotechar',
                    help="quotechar",
                    default = "'" )
    ap.add_argument('-t', '--headerrow',
                    help="headerrow; 0=row 1 (default = 0)",
                    default = 0 )
    args = vars(ap.parse_args())
    return args


def end_prog(code=0):
    stderr(f'einde: {datetime.today().strftime("%H:%M:%S")}')
    sys.exit(code)

 
if __name__ == "__main__":
    stderr(f'start: {datetime.today().strftime("%H:%M:%S")}')

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

