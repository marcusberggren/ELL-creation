import pandas as pd
import xlwings as xw
import numpy as np
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import fitz
import re
import os
import functions as fn

def main():
    data()
    return_pdf_df()


def get_pdf():
    root = tk.Tk()
    root.lift()
    root.withdraw()
    
    filename =  filedialog.askopenfilename(initialdir = xw.Book.caller(), title = "Select file", filetypes=[("Pdf files",".pdf")])
    root.quit()
    
    if filename:
        xw.Book.caller().sheets('START').range('E2').value = filename
    else:
        exit()

def get_ell():
    root = tk.Tk()
    root.lift()
    root.withdraw()
    
    filename =  filedialog.askopenfilename(initialdir = xw.Book.caller(), title = "Select file", filetypes=[("Excel files",".xlsx")])
    root.quit()
    
    if filename:
        xw.Book.caller().sheets('START').range('E5').value = filename
    else:    
        exit()

def data():
    file_path = xw.Book.caller().sheets('START').range('E2').value

    if file_path:
        doc = fitz.open(file_path)

        shipper = (33, 140)
        unit = (222, 290)
        goods = (312, 445)
        netw = (468, 502)
        tare = (510, 547)
        pod = (460, 500)
        base_triangle = (33, 100, 547, 580)

        total_starts = []  # lista över rektanglar för "Shipper :" på alla blad
        total_stops = []    # lista över rektanglar för "Units :" på alla blad
        total_height = 0.0  # loop över summan av höjderna på alla blad
        total_words = []    # lista över alla dokumentets ord på alla blad
        rect_pods = []     # lista över rektanglar för "Port of Discharge" på alla blad
        words_in_rectangles = [] # lista för ord inom rektang*larna, utan "headers"
        list_of_base_triangles = []
        list_of_rectangles = []
        list_of_tod_rectangles = []
        final_tod_list = []

        for page in doc:
            # starts
            starts = page.search_for("Shipper :")
            for rect in starts:
                total_starts.append(rect + (0, total_height, 0, total_height)) # lägger till på höjden för y0 och y1

            # stops
            stops = page.search_for("Units :")
            for rect in stops:
                total_stops.append(rect + (0, total_height, 0, total_height))

            # total_words
            words_in_page = page.get_text("words")
            for words in words_in_page:
                total_words.append([words[0], words[1] + total_height, words[2], words[3] + total_height, words[4]])

            # rect_pods
            pods = page.search_for('Port of discharge')
            for rect in pods:      
                rect_pods.append(rect + (0, total_height + 12, 0, total_height + 16)) # lägger till på y0 och y1 för att hitta POD: y0=12.20998, y1=16.08598

            # words_in_rectangles
            base_triangle = (33, 100 + total_height, 547, 580 + total_height) # Utgår från basrektangel (33, 100, 547, 580) och lägger på totalhöjden
            words_in_rectangles += [word for word in total_words if fitz.Rect(word[:4]).intersects(base_triangle)] # Tar fram alla ord inom rektangeln ovan
            list_of_base_triangles.append(base_triangle)

            total_height += page.rect.height # lägger till totalhöjden varje loop


        pod_names = [' '.join([words[4] for words in total_words if fitz.Rect(words[:4]).intersects(pod)]) for pod in rect_pods] # Tar fram alla POD is en lista

        rectangle = lambda x: fitz.Rect(33, total_starts[x][1], 580, total_stops[x][3])

        for num, starts in enumerate(total_starts):
            list_of_rectangles.append(rectangle(num))

        for rectangles, tod in zip(list_of_base_triangles, pod_names):
            list_of_tod_rectangles.append([rectangles[0], rectangles[1], rectangles[2], rectangles[3], tod])

        for rect in list_of_rectangles:
            counter = 0
            for tod in list_of_tod_rectangles:
                if tod[3] // rect[3] > 0 and counter == 0:
                    final_tod_list.append(tod[4])
                    counter == 1
                    break
                

        data.shipper = shipper
        data.unit = unit
        data.goods = goods
        data.netw = netw
        data.tare = tare
        data.pod = pod

        data.total_starts = total_starts
        data.total_stops = total_stops
        data.words_in_rectangles = words_in_rectangles
        data.list_of_rectangles = list_of_rectangles
        data.final_tod_list = final_tod_list


def get_container_nos(unit, start, stop):
    rectangle = (unit[0], start[1], unit[1], stop[3])
    same_row = start[1]
    list_of_data = []
    list_of_cont = []
    str_of_words = ""

    for word in data.words_in_rectangles:
        if fitz.Rect(word[:4]).intersects(rectangle):
            if same_row == word[1]:
                str_of_words += " " + word[4]

            elif same_row != word[1]:
                list_of_data.append(str_of_words.strip())
                str_of_words = word[4]

            same_row = word[1]

    for str in list_of_data:
        #matching CONTAINER
        container_match = re.search(r'^\w{4}\s\d{6}\-\d', str)
        trim_cont = re.sub(r'[\s-]', '', str)

        if container_match:
            list_of_cont.append(trim_cont)

    return list_of_cont

def get_shipper(shipper, start, stop):
    rectangle = (shipper[0], start[1], shipper[1], stop[3])
    same_row = start[1]
    list_of_data = []
    str_of_words = ""
    counter = 0

    for word in data.words_in_rectangles:
        if fitz.Rect(word[:4]).intersects(rectangle):
            
            if same_row == word[1]:                                
                str_of_words += " " + word[4]

            elif same_row != word[1]:                             
                list_of_data.append(str_of_words.strip())
                str_of_words = word[4]

            same_row = word[1]

    for str in list_of_data:
        #matching SHIPPER
        shipper_match = re.search('Shipper :', str)

        if counter == 1:
            counter = 0
            return str

        elif shipper_match:
            counter = 1

def get_netw(netw, start, stop):
    rectangle = (netw[0], start[1], netw[1], stop[3])
    list_of_data = []

    for word in data.words_in_rectangles:
        if fitz.Rect(word[:4]).intersects(rectangle):
            list_of_data.append(int(word[4]))

    if list_of_data:
        list_of_data.pop()

    return list_of_data

def get_tare(tare, start, stop):
    rectangle = (tare[0], start[1], tare[1], stop[3])
    list_of_data = []

    for word in data.words_in_rectangles:
        if fitz.Rect(word[:4]).intersects(rectangle):
                list_of_data.append(int(word[4]))

    if list_of_data:
        list_of_data.pop()

    return list_of_data

def get_netw_and_tare(netw, tare, start, stop):
    rect_netw = (netw[0], start[1], netw[1], stop[3])
    rect_tare = (tare[0], start[1], tare[1], stop[3])

    list_of_netw = []
    list_of_y1_netw = []
    list_of_tare = []
    list_of_y1_tare = []

    for word in data.words_in_rectangles:

        if fitz.Rect(word[:4]).intersects(rect_netw):
            list_of_netw.append(int(word[4]))
            list_of_y1_netw.append(int(word[3]))

        if fitz.Rect(word[:4]).intersects(rect_tare):
            list_of_tare.append(int(word[4]))
            list_of_y1_tare.append(int(word[3]))

    if list_of_netw:
        list_of_netw.pop()
        list_of_y1_netw.pop()
    
    if list_of_tare:
        list_of_tare.pop()
        list_of_y1_tare.pop()


    tare_list_diff = [x for x, y in enumerate(list_of_y1_netw) if y not in set(list_of_y1_tare)]
    netw_list_diff = [x for x in list_of_y1_tare if x not in set(list_of_y1_netw)]

    if tare_list_diff:
        for x in tare_list_diff:
            list_of_tare.insert(x, 0)

    if netw_list_diff:
        for x in netw_list_diff:
            list_of_netw.insert(x, 0)

    return list_of_netw, list_of_tare

def get_goods_info(goods, start, stop):
    rectangle = (goods[0], start[1], goods[1], stop[3])
    samma_rad = start
    lista_container_data = []
    lista_ord = ""

    for word in data.words_in_rectangles:

        if fitz.Rect(word[:4]).intersects(rectangle):

            if samma_rad == word[1]:                                # Om y-värdet på föregående rad i loop är samma som y-värdet för denna runda
                lista_ord += " " + word[4]

            elif samma_rad != word[1]:                              # Om y-värdet inte stämmer överens
                lista_container_data.append(lista_ord.strip())
                lista_ord = word[4]

            samma_rad = word[1]                                     # Sparar Y-värdet för raden för denna loop. Används i nästa runda

    lista_container_data.remove('')                                  # Tar bort första "" i listan
    
    customs_found = [id for id, value in enumerate(lista_container_data) if re.search(r'^Customs', value)]  # Söker efter "Customs" i lista_container_data och letar match
    
    lista_manifest_data = lista_container_data[customs_found[0]:]                                         # Ny lista som kopierar info från match med "Customs" och till slutet
    lista_container_data = lista_container_data[:customs_found[0]]                                            # Gör om lista och tar info fram till "Customs" match

    return lista_container_data, lista_manifest_data

def manifest_details(manifest_list):
    ocean_vessel_list = []
    final_pod_list = []
    counter_pod = 0
    counter_booking = 0
    counter_since_customs_match = 0
    voyage = ""

    for num, word in enumerate(manifest_list):

        #matching CUSTOMS STATUS
        customs_match = re.search(r'(?<=Customs status \")\w', word)
        if customs_match:
            customs_status = customs_match.group()
        else:
            counter_since_customs_match += 1

        #matching OCEAN VESSEL
        transhipment_match = re.search(r'^Transhipment by', word)
        voy_match = re.search(r'^.+(?= Voy)', word)
        vessel_match_voy = re.search(r'(?<=Transhipment by\s)(.+)([VvOoYy]{3})', word)
        vessel_match = re.search(r'(?<=Transhipment by\s).+[^\s*Voy\.*]', word)
        vessel_match_rest = re.search(r'^\w+[^ Voy]', word)

        if counter_since_customs_match < 5:
            if transhipment_match and vessel_match and voy_match:
                ocean_vessel_list.append(vessel_match_voy.group(1))

            elif transhipment_match and not voy_match:
                ocean_vessel_list.append(vessel_match.group())

            elif voy_match and not transhipment_match:
                ocean_vessel_list.append(vessel_match_rest.group())

        ocean_vessel = ' '.join(ocean_vessel_list).replace('Vessel ', '')

        #matching VOY
        voy_match = re.search(r'(?<=[VvOoYy]{3})(\s*\.*\s*)(\w+[-/]*\w*)', word)
        voy_match_last_row = re.search(r'[VvOoYy]{3}\s*\.*$', manifest_list[num-1])

        if counter_since_customs_match < 4:
            if voy_match:
                voyage = str(voy_match.group(2))
            elif voy_match_last_row:
                voyage = str(re.search(r'^\w+[-/]*\w*', word).group())
    

        #matching FINAL POD
        date_last_match = re.search(r'\d+\.+\d+\.+\d+$', word)
        etd_match = re.search(r'^ETD\s*\d+\.+\d+\.+\d+\s*\D+', word)
        only_etd_date = re.search(r'^(ETD\s*\d+\.+\d+\.+\d+)$', word)
        fpod_dual_dates = re.search(r'^(ETD\s*\d+\.+\d+\.+\d+\s*)(\D*)(\s+\d+\.+\d+\.+\d+)', word)
        fpod_single_date = re.search(r'^(ETD\s*\d+\.+\d+\.+\d+\s*)(\D*)', word)
        get_all_upper = re.search(r'^[\D]*', word)
        eta = re.search(r'ETA', word)
        eta_twice_same_row = re.search(r'ETA.+ETA', word)
        eta_twice_in_two_rows = re.search(r'ETA', manifest_list[num-1])

        def del_fillers_pod(string):
            new_string = re.sub(r'\s*ETA\s*', '', string)
            new_string2 = re.sub(r',[^,]*', '', new_string)
            new_string3 = re.sub(r'\s$', '', new_string2)
            return new_string3

        if counter_since_customs_match < 5:
            if etd_match and date_last_match:
                final_pod_list.append(del_fillers_pod(fpod_dual_dates.group(2)))

            elif etd_match and not date_last_match:
                final_pod_list.append(del_fillers_pod(fpod_single_date.group(2)))

            elif counter_pod == 1:
                final_pod_list.append(del_fillers_pod(get_all_upper.group()))
                counter_pod = 0

            elif counter_pod == 2:
                final_pod_list.append(del_fillers_pod(get_all_upper.group()))
                counter_pod == 0

            elif only_etd_date:
                counter_pod = 1

            elif etd_match and not date_last_match:
                counter_pod = 2

            elif eta and eta_twice_in_two_rows:
                final_pod_list.append(re.search(r'^(ETA\s*\d+\.+\d+\.+\d+\s*)(\D*)', manifest_list[num-1]).group(2))
            elif eta_twice_same_row:
                final_pod_list.append(re.search(r'^(ETA\s*\d+\.+\d+\.+\d+\s*)(\D*)(?=\sETA)', word).group(2))


        final_pod = ' '.join(final_pod_list)

        #matching BOOKING NUMBER
        ref_match = re.search(r'^ref', word, flags=re.IGNORECASE)

        def del_fillers_ref(string):
            new_string = re.sub(r'^[VvTtGg]*\s*[RrEeFf]{3}[:\.\s]*', '', string)
            new_string2 = re.sub(r'\s*OPS.*', '', new_string)
            return new_string2

        if ref_match and counter_booking != 1:
            booking_number = del_fillers_ref(word)

        elif counter_booking == 1:
            if transhipment_match:
                counter_booking = 0

            if not transhipment_match:
                booking_number = del_fillers_ref(word)
                counter_booking = 0

        elif customs_match:
            counter_booking = 1

    return_dict = {
        "BOOKING NUMBER": booking_number,
        "CUSTOMS STATUS": customs_status,
        "OCEAN VESSEL": ocean_vessel,
        "VOY" : voyage,
        "FINAL POD": final_pod
        }
    
    return return_dict

def goods_details(goods_list):
    list_of_unit_types = []
    list_of_packages = []
    list_of_goods = []
    list_of_load_status = []
    
    for num, val in enumerate(goods_list):

        match_unit_type = re.search(r"^\d{2}\'\D+", val)
        match_unit_type_last_row = re.search(r"^\d{2}\'\D+", goods_list[num-1])
        match_goods_info = re.search(r'^(\d+)\s+\w*\s+(\D+)', val)
        match_goods_info2 = re.search(r'^\D*$', val)
        match_empty = re.search(r'EMPTY|^MT', val)

        if match_unit_type:
            list_of_unit_types.append(match_unit_type.group())
        
        if match_goods_info and match_unit_type_last_row:
            list_of_packages.append(int(match_goods_info.group(1)))
            list_of_goods.append(match_goods_info.group(2))

        elif match_goods_info2 and match_unit_type_last_row:
            list_of_goods.append(match_goods_info2.group())


        if match_empty:
            list_of_load_status.append("MT")
        elif match_goods_info and match_unit_type_last_row:
            list_of_load_status.append("LA")
        elif match_goods_info2 and match_unit_type_last_row:
            list_of_load_status.append("LA")


    return list_of_unit_types, list_of_packages, list_of_goods, list_of_load_status

def cargo_detail():

    file_path = file_path = xw.Book.caller().sheets('START').range('E5').value
    
    if file_path:
        
        with xw.App(visible=False) as app:
            wb = app.books.open(file_path)
            sheet = wb.sheets('Cargo Detail')
            last_row = sheet.range('AD' + str(sheet.cells.last_cell.row)).end('up').row
            rng_cargo_detail = sheet.range('A5:BF' + str(last_row))

            cargo_detail.vessel = sheet.range('A2').value
            cargo_detail.voy = sheet.range('B2').value
            cargo_detail.pol = sheet.range('F2').value
            cargo_detail.leg = sheet.range('C2').value

            df = sheet.range(rng_cargo_detail).options(pd.DataFrame, index=False, header=True).value
            df = pd.DataFrame(df).copy()

            cargo_detail.len_df = len(df)
            
            wb.close()

        cd_df = pd.DataFrame()

        cd_df.loc[:, 'CD CONTAINER'] = df['Container No']
        cd_df.loc[:,'CD TOD'] = df['Pod terminal']
        cd_df.loc[:, 'CD ISO TYPE'] = df['ISO Container Type']
        cd_df.loc[:, 'CD LOAD STATUS'] = df['Commodity']
        cd_df.loc[:, 'CD GWT'] = df['Weight in MT']
        cd_df.loc[:, 'IMDG'] = df['IMCO']
        cd_df.loc[:, 'UNNR'] = df['UN']
        cd_df.loc[:, 'TEMP'] = df['TempOpt']
        cd_df.loc[:, 'OOG'] = df['OOG']
        cd_df.loc[:, 'REMARK'] = df['Remarks']

        return cd_df

def run_it_all():

    df_cd = cargo_detail().copy()
    bookings_count = len(data.list_of_rectangles)

    list_of_booking_refs = []
    list_of_shipper = []
    list_of_tods = []
    list_of_containers = []
    list_of_container_types = []
    list_of_netw = []
    list_of_tare = []
    list_of_load_status = []
    list_of_customs_status = []
    list_of_packages = []
    list_of_goods = []
    list_of_ocean_vessels = []
    list_of_voy = []
    list_of_final_pod = []
    
    for booking_no in range(bookings_count):
        #setup
        start = data.total_starts[booking_no]
        stop = data.total_stops[booking_no]

        container_list = get_container_nos(data.unit, start, stop)
        package_list = goods_details(get_goods_info(data.goods, start, stop)[0])[1]
        booking_ref_list = [manifest_details(get_goods_info(data.goods, start, stop)[1])["BOOKING NUMBER"]] * len(container_list)
        netw_list = get_netw_and_tare(data.netw, data.tare, start, stop)[0]
        tare_list = get_netw_and_tare(data.netw, data.tare, start, stop)[1]
        voy_list = [manifest_details(get_goods_info(data.goods, start, stop)[1])["VOY"]] * len(container_list)

        #creation of lists
        list_of_shipper += [get_shipper(data.shipper, start, stop)]*len(container_list)
        list_of_tods += [data.final_tod_list[booking_no]] * len(container_list)
        list_of_containers += container_list
        list_of_container_types += goods_details(get_goods_info(data.goods, start, stop)[0])[0]

        if not package_list:
            list_of_packages += [0] * len(container_list)
        else:
            list_of_packages += package_list
            
        list_of_goods += goods_details(get_goods_info(data.goods, start, stop)[0])[2]
        list_of_customs_status += manifest_details(get_goods_info(data.goods, start, stop)[1])["CUSTOMS STATUS"] * len(container_list)
        list_of_ocean_vessels += [manifest_details(get_goods_info(data.goods, start, stop)[1])["OCEAN VESSEL"]] * len(container_list)
        list_of_final_pod += [manifest_details(get_goods_info(data.goods, start, stop)[1])["FINAL POD"]] * len(container_list)
        list_of_load_status += goods_details(get_goods_info(data.goods, start, stop)[0])[3]

        if not voy_list:
            list_of_voy += [""] * len(container_list)
        else:
            list_of_voy += voy_list

        if not booking_ref_list:
            list_of_booking_refs += [""] * len(container_list)
        else:
            list_of_booking_refs += booking_ref_list

        if not netw_list and not tare_list:
            list_of_netw += [0] * len(container_list)
            list_of_tare += [0] * len(container_list)
        else:
            list_of_netw += netw_list
            list_of_tare += tare_list
    
    list_of_pods = [a.title() for a in list_of_final_pod]
    
    cd_dict = {}
    cd_list_of_tod = []
    cd_list_of_iso_type = []
    cd_list_of_load_status = []
    cd_list_of_gwt = []
    cd_list_of_imdg = []
    cd_list_of_unnr = []
    cd_list_of_temp = []
    cd_list_of_oog = []
    cd_list_of_remark = []

    for val in df_cd.values:
        cd_dict[val[0]] = [val[1], val[2], val[3], val[4], val[5], val[6], val[7], val[8], val[9]]

    for val in list_of_containers:
        if val in cd_dict.keys():
            cd_list_of_tod.append(cd_dict[val][0])
            cd_list_of_iso_type.append(cd_dict[val][1])
            cd_list_of_load_status.append(cd_dict[val][2])
            cd_list_of_gwt.append(cd_dict[val][3])
            cd_list_of_imdg.append(cd_dict[val][4])
            cd_list_of_unnr.append(cd_dict[val][5])
            cd_list_of_temp.append(cd_dict[val][6])
            cd_list_of_oog.append(cd_dict[val][7])
            cd_list_of_remark.append(cd_dict[val][8])
        else:
            cd_list_of_tod.append(np.nan)
            cd_list_of_iso_type.append(np.nan)
            cd_list_of_load_status.append(np.nan)
            cd_list_of_gwt.append(np.nan)
            cd_list_of_imdg.append(np.nan)
            cd_list_of_unnr.append(np.nan)
            cd_list_of_temp.append(np.nan)
            cd_list_of_oog.append(np.nan)
            cd_list_of_remark.append(np.nan)

    l1 = []
    
    for key, val in cd_dict.items():
        if key not in list_of_containers:
            l1.append(['CARGO DETAIL', key, val[0], val[1], val[2], val[3], val[4], val[5], val[6], val[7], val[8]])
        arr1 = np.array(l1)

    if len(arr1) > 0:
        cd_df_rest = pd.DataFrame(arr1, columns=['MATCH', 'CONTAINER', 'CD TOD', 'CD ISO TYPE', 'CD LOAD STATUS', 'GWT', 'IMDG', 'UNNR', 'TEMP', 'OOG', 'REMARK']).fillna('')


    dict_final = {
        'MATCH': "",
        'BOOKING NUMBER': list_of_booking_refs,
        'MLO': list_of_shipper,
        'MAN TOD': list_of_tods,
        'CD TOD': cd_list_of_tod,
        'CONTAINER': list_of_containers,
        'MAN ISO TYPE': list_of_container_types,
        'CD ISO TYPE': cd_list_of_iso_type,
        'NET WEIGHT': list_of_netw,
        'MAN LOAD STATUS': list_of_load_status,
        'CD LOAD STATUS': cd_list_of_load_status,
        'GWT': cd_list_of_gwt,
        'IMDG': cd_list_of_imdg,
        'UNNR': cd_list_of_unnr,
        'TEMP': cd_list_of_temp,
        'OOG': cd_list_of_oog,
        'REMARK': cd_list_of_remark,
        'TARE': list_of_tare,
        'MLO PO': "",
        'CUSTOMS STATUS': list_of_customs_status,
        'PACKAGES': list_of_packages,
        'GOODS DESCRIPTION': list_of_goods,
        'OCEAN VESSEL': list_of_ocean_vessels,
        'VOY' : list_of_voy,
        'FINAL POD': list_of_pods
        }

    #dict_final = {
    #    'MATCH': "",
    #    'BOOKING NUMBER': "",
    #    'MLO': "",
    #    'MAN TOD': "",
    #    'CD TOD': "",
    #    'CONTAINER': "",
    #    'MAN ISO TYPE': "",
    #    'CD ISO TYPE': "",
    #    'NET WEIGHT': "",
    #    'MAN LOAD STATUS': list_of_load_status,
    #    'CD LOAD STATUS': "",
    #    'GWT': "",
    #    'IMDG': "",
    #    'UNNR': "",
    #    'TEMP': "",
    #    'OOG': "",
    #    'REMARK': "",
    #    'TARE': "",
    #    'MLO PO': "",
    #    'CUSTOMS STATUS': "",
    #    'PACKAGES': "",
    #    'GOODS DESCRIPTION': "",
    #    'OCEAN VESSEL': "",
    #    'VOY': "",
    #    'FINAL POD': ""
    #    }   

    df = pd.DataFrame(dict_final)
    df.loc[df['CD TOD'].notna(), 'MATCH'] = "MATCH"
    df.loc[df['CD TOD'].isna(), 'MATCH'] = "MANIFEST"
    df.loc[(df['MAN ISO TYPE'].str[0] == "2") & (df['TARE'] == 0), 'TARE'] = 2200
    df.loc[(df['MAN ISO TYPE'].str[0] == "4") & (df['TARE'] == 0), 'TARE'] = 4000
    df.loc[(df['MAN ISO TYPE'].str == "22T1") & (df['TARE'] == 0), 'TARE'] = 7000

    run_it_all.len_df = len(df)

    if len(arr1) > 0:
        df_concat = pd.concat([df, cd_df_rest], ignore_index=True, join='outer')
    else:
        df_concat = df.copy()

    return(df_concat)

def clean_data_output():
    df = run_it_all()
    df = fn.regex_no_extra_whitespace(df).copy()

    sheet = xw.Book.caller().sheets('LIB')

    lib_range = sheet.range('A2').expand()
    lib_dict = sheet.range(lib_range).options(dict).value

    df.loc[:, 'MLO'] = df['MLO'].replace(lib_dict)
    df.loc[:, 'MAN TOD'] = fn.get_template_type(df, ['tpl_terminal', 'TERMINAL OUTPUT', 'MAN TOD'])
    df.loc[:, 'MAN ISO TYPE'] = fn.get_template_type(df, ['tpl_cargo_type', 'TYPE OUTPUT', 'MAN ISO TYPE'])
    df.loc[:, 'FINAL POD'] = fn.get_template_type(df, ['tpl_ports', 'UNLOCODE', 'FINAL POD'])
    
    list_of_po = []

    for i in range(len(df)):
        shipper = df.iloc[i, 2] # 1 symboliserar 'MLO'
        if isinstance(shipper, float):
            list_of_po.append(df.iloc[i, 1]) # 0 symboliserar 'BOOKING NUMBER'
        elif 'HL' in shipper:
            list_of_po.append("")
        elif 'ONE' in shipper:
            list_of_po.append("")
        else:
            list_of_po.append(df.iloc[i, 1]) # 0 symboliserar 'BOOKING NUMBER'

    df['MLO PO'] = list_of_po.copy()

    return df

def return_pdf_df():
    wb = xw.Book.caller()
    sheet = wb.sheets('RESULT')
    sheet.range('A5').options(pd.DataFrame, header=False).value = clean_data_output().copy()

    sheet.range('C1').value = run_it_all.len_df
    sheet.range('C2').value = cargo_detail.len_df
    sheet.range('F1').value = cargo_detail.vessel
    sheet.range('G1').value = cargo_detail.voy
    sheet.range('H1').value = cargo_detail.pol
    sheet.range('I1').value = cargo_detail.leg

def prep_ell_data():
    wb = xw.Book.caller()
    sheet = wb.sheets('RESULT')
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    data_range = sheet.range('A5:Z' + str(last_row))
    data = sheet.range(data_range).options(pd.DataFrame, index=False, header=False).value
    df = pd.DataFrame(data)
    df = df.rename(columns= {
        0: '',
        1: 'MATCH',
        2: 'BOOKING NUMBER',
        3: 'MLO',
        4: 'MAN TOD',
        5: 'CD TOD',
        6: 'CONTAINER',
        7: 'MAN ISO TYPE',
        8: 'CD ISO TYPE',
        9: 'NET WEIGHT',
        10: 'MAN LOAD STATUS',
        11: 'CD LOAD STATUS',
        12: 'GWT',
        13: 'IMDG',
        14: 'UNNR',
        15: 'TEMP',
        16: 'OOG',
        17: 'REMARK',
        18: 'TARE',
        19: 'MLO PO',
        20: 'CUSTOMS STATUS',
        21: 'PACKAGES',
        22: 'GOODS DESCRIPTION',
        23: 'OCEAN VESSEL',
        24: 'VOY',
        25: 'FINAL POD'
        })

    df_pod = fn.get_csv_data('terminal').copy()
    dict_pod = df_pod.set_index('TERMINAL')['PORT'].to_dict()

    df_cargo_type = fn.get_csv_data('cargo_type').copy()
    dict_ct = df_cargo_type.set_index('ISO STATUS')['T6 CARGO TYPE'].to_dict()

    def get_cd_df(df: pd.DataFrame):
        df_cd = pd.DataFrame(columns=['Pod', 'Pod call seq', 'Pod terminal', 'Pod Status', 'POL',
        'Pol terminal', 'Pol Status', 'Shunted Terminal', 'Slot Owner', 'Slot Account', 'MLO', 'MLO PO',
        'Booking Reference', 'Ex Vessel', 'Ex Voyage', 'Next Vessel', 'Next Voyage', 'Mother Vessel',
        'Mother Vessel CallSign', 'Mother Voyage', 'POT', 'F.Destination', 'VIA', 'VIA terminal', 'Cargo type',
        'ISO Container Type', 'User Container Type', 'Commodity', 'OOG', 'Container No', 'Weight in MT', 'Stowage',
        'Door Open', 'Slot Killed', 'V.Type', 'Fr.Group', 'TempMax', 'TempMin', 'TempOpt', 'IMCO', 'FP', 'UN',
        'PSA Class', 'IMO Name', 'Chem Name', 'Remarks', 'OOH(CM)', 'OLF(CM)', 'OLA(CM)', 'OWP(CM)', 'OWS(CM)',
        'VGM Weight in MT', 'VGM Cert Signatory', 'VGM Certificate No', 'VGM Weighing Method',
        'VGM Cert Issuing Party', 'VGM Cert Issuing Address', 'VGM Cert Issue Date'])

        df.loc[:, 'CONCAT'] = df['MAN ISO TYPE'] + df['MAN LOAD STATUS']

        sheet = xw.Book.caller().sheets('LIB')
        lib_range = sheet.range('D2').expand()
        lib_dg_dict = sheet.range(lib_range).options(dict).value

        df_cd.loc[:, 'Pod'] = df['MAN TOD']
        df_cd.loc[:, 'Pod'] = df['MAN TOD'].replace(dict_pod)
        df_cd.loc[:, 'Pod call seq'] = 1
        df_cd.loc[:, 'Pod terminal'] = df['MAN TOD']
        df_cd.loc[df['MAN LOAD STATUS'] == 'MT', 'Pod Status'] = "L"
        df_cd.loc[df['MAN LOAD STATUS'] == 'LA', 'Pod Status'] = "T"
        df_cd.loc[:, 'POL'] = 'FITOR'
        df_cd.loc[:, 'Pol terminal'] = 'FIOUK'
        df_cd.loc[:, 'Pol Status'] = "L"
        df_cd.loc[:, 'Slot Owner'] = "XCL"
        df_cd.loc[:, 'Slot Account'] = "XCL"
        df_cd.loc[:, 'MLO'] = df['MLO']
        df_cd.loc[:, 'MLO PO'] = df['MLO PO']
        df_cd.loc[:, 'Booking Reference'] = df['BOOKING NUMBER']
        df_cd.loc[:, 'Mother Vessel'] = df['OCEAN VESSEL']
        df_cd.loc[:, 'Mother Vessel CallSign'] = np.nan
        df_cd.loc[:, 'Mother Voyage'] = df['VOY']
        df_cd.loc[:, 'F.Destination'] = df['FINAL POD']
        df_cd.loc[:, 'Cargo type'] = df['CONCAT'].replace(dict_ct)
        df_cd.loc[:, 'ISO Container Type'] = df['MAN ISO TYPE']
        df_cd.loc[:, 'User Container Type'] = df['MAN ISO TYPE']
        df_cd.loc[:, 'Commodity'] = df['MAN LOAD STATUS']
        df_cd.loc[:, 'Container No'] = df['CONTAINER']
        df_cd.loc[:, 'Weight in MT'] = df['GWT']
        df_cd.loc[:, 'TempMax'] = df['TEMP']
        df_cd.loc[:, 'TempMin'] = df['TEMP']
        df_cd.loc[:, 'TempOpt'] = df['TEMP']
        df_cd.loc[:, 'IMCO'] = df['IMDG']
        df_cd.loc[df_cd['IMCO'].notna(), 'Cargo type'] = df['CONCAT'].replace(lib_dg_dict) #Cargo type again
        df_cd.loc[:, 'UN'] = df['UNNR']
        df_cd.loc[df_cd['IMCO'].notna(), 'IMO Name'] = "CHEM"
        df_cd.loc[:, 'Remarks'] = df['REMARK']
        df_cd.loc[df['MAN LOAD STATUS'] != "MT", 'VGM Weight in MT'] = df['GWT']
        return df_cd
    

    def get_man_df(df: pd.DataFrame):
        df_man = pd.DataFrame(columns=['Pod Terminal', 'MLO', 'B/L No', 'MLO PO', 'Booking Reference',
        'OBL Reference', 'Marks & Nos', 'No of Cntr', 'Type', 'Stc', 'No of Packages', 'Unit', 'Goods Desc',
        'Cargo Status', 'Transhipment', 'Seal No', 'Tare Weight in Kilos', 'Net Weight in Kilos',
        'Deep Sea Vessel', 'ETA', 'Rcvr', 'Shipper', 'Consignee', 'Notify', 'Product ID', 'Volume (Meter)',
        'Marks', 'MRN Number', 'Remarks', 'Port of Origin', 'Export PO', 'Package Content'])

        df.loc[:, 'POD SHORT'] = df['MAN TOD'].str[:2]
        df.loc[:, 'POD MANIFEST'] = fn.get_template_type(df, ['country', 'COUNTRY', 'POD SHORT'])

        df_man.loc[:, 'Pod Terminal'] = df['MAN TOD']
        df_man.loc[:, 'Pod Terminal'] = df['MAN TOD'].replace(dict_pod)
        df_man.loc[:, 'MLO'] = df['MLO']
        df_man.loc[:, 'MLO PO'] = df['MLO PO']
        df_man.loc[:, 'Booking Reference'] = df['BOOKING NUMBER']
        df_man.loc[:, 'Marks & Nos'] = df['CONTAINER']
        df_man.loc[:, 'No of Cntr'] = 1
        df_man.loc[:, 'Type'] = df['MAN ISO TYPE']
        df_man.loc[:, 'Stc'] = "STC"
        df_man.loc[:, 'No of Packages'] = df['PACKAGES']
        df_man.loc[:, 'Unit'] = "PK"
        df_man.loc[:, 'Goods Desc'] = df['GOODS DESCRIPTION']
        df_man.loc[:, 'Cargo Status'] = df['CUSTOMS STATUS']
        df_man.loc[df['MAN LOAD STATUS'] == 'MT', 'Transhipment'] = "N"
        df_man.loc[df['MAN LOAD STATUS'] == 'LA', 'Transhipment'] = "Y"
        df_man.loc[:, 'Tare Weight in Kilos'] = df['TARE']
        df_man.loc[:, 'Net Weight in Kilos'] = df['NET WEIGHT']
        df_man.loc[:, 'Deep Sea Vessel'] = df['OCEAN VESSEL']
        df_man.loc[:, 'ETA'] = ""
        df_man.loc[:, 'Shipper'] = df['MLO'] + ' FINLAND'
        df_man.loc[:, 'Consignee'] = df['MLO'] + " " + df['POD MANIFEST']
        df_man.loc[:, 'Notify'] = df_man['Consignee']
        df_man.loc[:, 'MRN Number'] = ""
        df_man.loc[:, 'Package Content'] = df['GOODS DESCRIPTION']
        return df_man

    return get_cd_df(df), get_man_df(df)


def export_ell_data():

    df1, df2 = prep_ell_data()

    wb = xw.Book.caller()
    sheet = wb.sheets('RESULT')
    wb_path = wb.sheets('START').range('E5').value

    vessel = sheet.range('F1').value
    voy = str(sheet.range('G1').value) #sparar som sträng
    voyage = re.search(r'^\d{0,5}', voy).group(0) #tar bara fram de 5 första siffrorna
    leg = sheet.range('I1').value
    pol = sheet.range('H1').value

    folder_path_ell = os.path.split(wb_path)[0]
    time_str = datetime.now().strftime("%y%m%d")
    ell_file_name = "ELL_" + vessel + "_" + str(voyage) + "_" + pol + "_" + time_str + ".xlsx"
    name_of_file_and_path = os.path.join(folder_path_ell, ell_file_name)
    
    with xw.App(visible=False) as app:
        wb = app.books.open(fn.get_path('tpl_ell'))
        wb.save(name_of_file_and_path)
        cargo_detail_sheet = wb.sheets['Cargo Detail']
        manifest_sheet = wb.sheets['Manifest']
        cargo_detail_sheet.range('A6').options(pd.DataFrame, index=False, header=False).value = df1.copy()
        cargo_detail_sheet.range('A2').value = vessel
        cargo_detail_sheet.range('B2').value = voyage
        cargo_detail_sheet.range('C2').value = leg
        cargo_detail_sheet.range('F2').value = pol
        manifest_sheet.range('A2').options(pd.DataFrame, index=False, header=False).value = df2.copy()
        wb.save()
        wb.close()

#if __name__ == "__main__":
    #file_path = r'C:\Users\SWV224\BOLLORE\XPF - Documents\MAINTENANCE\Test files\FITOR_Makro.xlsm'
    #xw.Book(file_path).set_mock_caller()
