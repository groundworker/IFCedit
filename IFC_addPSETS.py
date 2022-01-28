# -*- coding: utf-8 -*-
"""
@author: Andreas Sorgatz-Wenzel
"""

import re
import random
import string
import pandas
from pathlib import Path


def read_ascii(filename):
    file_in = open(filename, 'r')
    file_in_inhalt = file_in.read()
    file_in.close()
    file_in = open(filename, 'r')
    inhalt = file_in.readlines()
    file_in.close()
    file_out = open(filename[:-4]+"_output.ifc", 'w')
    return inhalt, file_out, file_in_inhalt


def read_excel(filename, sheet):
    dataframe = pandas.read_excel(filename, sheet_name=sheet)
    return dataframe


def randomString(stringLength=22):
    lettersAndDigits = string.ascii_letters + string.digits
    rS = ''.join(random.choice(lettersAndDigits) for i in range(stringLength))
    return rS


def addIFCProp(n_new, file_out, path, file_excel, excel_sheet):
    check_excel = False
    try:
        df = read_excel(path+file_excel, excel_sheet)
        print(str(len(df['Name']))+' Objekte in Excel-Tabelle ('+excel_sheet+') gefunden')
        check_excel = True
    except:
        print('Excel-Datei ('+str(file_excel)+') oder Excel-Tabelle ('+str(excel_sheet)+') nicht vorhanden')
    if check_excel:
        try:
            for index, name in enumerate(df['Name']):
                print(str(name)+' - wird gesucht')
                check = False
                for element in list_element:
                    if str(name) == element.Name[1:-1]:
                        print(str(name)+' - konnte zugeordnet werden')
                        n = n_new
                        counter = 0
                        number_values = ''
                        for i in range(2, len(df.columns)):
                            counter += 1
                            number_values += '#'+str(n+counter)+','
                            file_out.write("#"+str(n+counter)+"=IFCPROPERTYSINGLEVALUE('"+df.columns[i]+"',$,IFCLABEL('"+str(df[df.columns[i]][index])+"'),$);\n")
                        file_out.write("#"+str(n+counter+1)+"=IFCPROPERTYSET('"+str(randomString(22))+"',$,'"+excel_sheet+"',$,("+number_values[:-1]+"));\n")
                        file_out.write("#"+str(n+counter+2)+"=IFCRELDEFINESBYPROPERTIES('"+str(randomString(22))+"',$,$,$,("+str(element.Number)+"),#"+str(n+counter+1)+");\n")
                        n_new = n+counter+2
                        check = True
                        break
                    elif df['GUID'][index] == element.GlobalId[1:-1]:
                        print('GUID ('+df['GUID'][index] + ') von ' + df['Name'][index] +
                              ' - konnte zugeordnet werden')
                        n = n_new
                        counter = 0
                        number_values = ''
                        for i in range(2, len(df.columns)):
                            counter += 1
                            number_values += '#'+str(n+counter)+','
                            file_out.write("#"+str(n+counter)+"=IFCPROPERTYSINGLEVALUE('"+df.columns[i]+"',$,IFCLABEL('"+str(df[df.columns[i]][index])+"'),$);\n")
                        file_out.write("#"+str(n+counter+1)+"=IFCPROPERTYSET('"+str(randomString(22))+"',$,'"+excel_sheet+"',$,("+number_values[:-1]+"));\n")
                        file_out.write("#"+str(n+counter+2)+"=IFCRELDEFINESBYPROPERTIES('"+str(randomString(22))+"',$,$,$,("+str(element.Number)+"),#"+str(n+counter+1)+");\n")
                        n_new = n+counter+2
                        check = True
                        break
                if check is False:
                    print(str(name)+' - konnte nicht zugeordnet werden')
        except Exception as e:
            print(e)
            print('Zuordnung nicht möglich')
    return n_new, file_out


class IFCELEMENT:
    def __init__(self, number, ifcproduct, ifcpara):
        # number (location) of product in IFC-file 
        self.Number = number
        # kind of IFC-product
        self.Product = ifcproduct
        # official IFC-parameter for any subtype of IFCELEMENT
        self.GlobalId = ifcpara[0]          # 1
        self.OwnerHistory = ifcpara[1]      # 2
        self.Name = ifcpara[2]              # 3
        self.Description = ifcpara[3]       # 4
        self.ObjectType = ifcpara[4]        # 5
        self.ObjectPlacement = ifcpara[5]   # 6
        self.Representation = ifcpara[6]    # 7
        self.Tag = ifcpara[7]               # 8


if __name__ == '__main__':
    # input
    file = 'geological-model.ifc'
    ifc_element_type = 'IFCBUILDINGELEMENTPROXY'
    file_excel = 'Merkmale_geological-model.xlsx'
    excel_sheet = ['Pset_A','Pset_B']

    path = str(Path().absolute())+'\\'
    
    # open ifc file
    content, file_out, file_in_inhalt = read_ascii(path+file)
    ifccontent = []
    list_element = []
    content_typ = ''

    # search for ifcelement
    for line in content:
        line = line[:-2]
        if line == 'HEADER':
            content_typ = 'header'
        elif line == 'ENDSEC' and content_typ == 'header':
            content_typ = 'end_header'
        elif line == 'DATA':
            content_typ = 'data'
        elif line == 'ENDSEC' and content_typ == 'data':
            content_typ = 'end_data'
        if content_typ == 'data' and line != 'DATA':
            line_content = line.split('=')
            ifccontent.append(line_content)
            if ifc_element_type in line_content[1]:
                text = line_content[1][:-1]
                text = text.split('(')
                ifcproduct = text[0]
                PATTERN = re.compile(r'''((?:[^,"']|"[^"]*"|'[^']*')+)''')
                ifcpara = PATTERN.split(text[1])[1::2]
                list_element.append(IFCELEMENT(line_content[0], ifcproduct, ifcpara))
        if content_typ == 'end_data':
            print('IFC Datei vollständig eingelesen')
            break
    print(str(len(list_element))+' Objekte ('+ifc_element_type+') im Modell gefunden')
    number_last = int(ifccontent[(len(ifccontent)-1)][0].replace('#', ''))
    number_new = number_last
    file_out.write(file_in_inhalt[0:-27])
    file_out.write("\n")
    
    # comparison of model and excel sheet
    for sheet_element in excel_sheet:
        number_new, file_out = addIFCProp(number_new, file_out, path, file_excel, sheet_element)

    file_out.write("ENDSEC;\n")
    file_out.write("END-ISO-10303-21;")
    file_out.close()
