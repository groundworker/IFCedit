# -*- coding: utf-8 -*-
"""
@author: Andreas Sorgatz-Wenzel
"""

import re
import random
import string
import pandas
from pathlib import Path
import IfcDefinitions as IfcDef


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


def addIFCProp(n_new, file_out, path, file_excel, excel_sheet, content_ifcfile):
    check_excel = False
    try:
        df = read_excel(path+file_excel, excel_sheet)
        print(str(len(df['Suche']))+' Objekte in Excel-Tabelle ('+excel_sheet+') gefunden')
        check_excel = True
    except Exception as e:
        print(e)
        print('Excel-Datei ('+str(file_excel)+') oder Excel-Tabelle ('+str(excel_sheet)+') nicht vorhanden')
    for index, item in enumerate(df['Suche']):
        if item == 'object' and check_excel == True:
            print('Suche anhand der Objekte')
            try:
                for element in content_ifcfile.LISTELEMENT:
                    print(str(df['Name'][index])+' - wird gesucht')
                    check = False
                    if str(df['Name'][index]) == element.Name[1:-1] :
                        print(str(df['Name'][index])+' - konnte zugeordnet werden')
                        n = n_new
                        counter = 0
                        number_values = ''
                        for i in range(6, len(df.columns)):
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
                        for i in range(6, len(df.columns)):
                            counter += 1
                            number_values += '#'+str(n+counter)+','
                            file_out.write("#"+str(n+counter)+"=IFCPROPERTYSINGLEVALUE('"+df.columns[i]+"',$,IFCLABEL('"+str(df[df.columns[i]][index])+"'),$);\n")
                        file_out.write("#"+str(n+counter+1)+"=IFCPROPERTYSET('"+str(randomString(22))+"',$,'"+excel_sheet+"',$,("+number_values[:-1]+"));\n")
                        file_out.write("#"+str(n+counter+2)+"=IFCRELDEFINESBYPROPERTIES('"+str(randomString(22))+"',$,$,$,("+str(element.Number)+"),#"+str(n+counter+1)+");\n")
                        n_new = n+counter+2
                        check = True
                        break
                if check is False:
                    print(str(df['Name'][index])+' - konnte nicht zugeordnet werden')
            except Exception as e:
                print(e)
                print('Zuordnung nicht möglich')
        if check_excel == True and item == 'property':
            print('Suche anhand der Properties')
            try:
                if df['Suche'][index] == 'property':
                    print(str(df['Property'][index])+':'+str(df['Value'][index])+' - wird gesucht')
                    check = False
                    for ifcprop in content_ifcfile.LISTPROPERTYSINGLEVALUE:
                        if str(df['Property'][index]) in ifcprop.Name and df['Value'][index] in ifcprop.NominalValue:
                            for ifcpset in content_ifcfile.LISTPROPERTYSET:
                                if ifcprop.Number in ifcpset.HasProperties:
                                    for ifcproprel in content_ifcfile.LISTRELDEFINESBYPROPERTIES:
                                        if ifcpset.Number in ifcproprel.RelatingPropertyDefinition:
                                            object_number = ifcproprel.RelatedObjects
                                            print(str(df['Property'][index])+':'+str(df['Value'][index])+' in IFCELEMENT '+str(object_number)+' gefunden')
                                            break
                            n = n_new
                            counter = 0
                            number_values = ''
                            for i in range(6, len(df.columns)):
                                counter += 1
                                number_values += '#'+str(n+counter)+','
                                file_out.write("#"+str(n+counter)+"=IFCPROPERTYSINGLEVALUE('"+df.columns[i]+"',$,IFCLABEL('"+str(df[df.columns[i]][index])+"'),$);\n")
                            text_object_number = ''
                            for obj in object_number:
                                text_object_number += str(obj)+','
                            file_out.write("#"+str(n+counter+1)+"=IFCPROPERTYSET('"+str(randomString(22))+"',$,'"+excel_sheet+"',$,("+number_values[:-1]+"));\n")
                            file_out.write("#"+str(n+counter+2)+"=IFCRELDEFINESBYPROPERTIES('"+str(randomString(22))+"',$,$,$,("+text_object_number[:-1]+"),#"+str(n+counter+1)+");\n")
                            n_new = n+counter+2
                            check = True
                if check is False:
                    print(str(df['Property'][index])+' - konnte nicht zugeordnet werden')
            except Exception as e:
                print(e)
                print('Zuordnung nicht möglich')
    return n_new, file_out


class IFCFILE:
    def __init__(self, list_element, list_propertyset, list_reldefinesbyproperties, list_propertysinglevalue):
        self.LISTELEMENT = list_element
        self.LISTPROPERTYSET = list_propertyset
        self.LISTRELDEFINESBYPROPERTIES = list_reldefinesbyproperties
        self.LISTPROPERTYSINGLEVALUE = list_propertysinglevalue


if __name__ == '__main__':
    # input
    file = 'Altaufschluesse.ifc'
    ifc_element_type = 'IFCBUILDINGELEMENTPROXY'
    file_excel = 'Merkmale_IFC_addPSETS_searchbyPset.xlsx'
    excel_sheet = ['Pset_A','Pset_B']

    path = str(Path().absolute())+'\\'
    
    # open ifc file
    content, file_out, file_in_inhalt = read_ascii(path+file)
    ifccontent = []
    list_element = []
    list_propertyset = []
    list_reldefinesbyproperties = []
    list_propertysinglevalue = []
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
            line_content = line.split('=',1)
            ifccontent.append(line_content)
            if ifc_element_type in line_content[1]:
                text = line_content[1][:-1]
                text = text.split('(')
                ifcproduct = text[0]
                PATTERN = re.compile(r'''((?:[^,"']|"[^"]*"|'[^']*')+)''')
                ifcpara = PATTERN.split(text[1])[1::2]
                list_element.append(IfcDef.IFCELEMENT(line_content[0], ifcproduct, ifcpara))
            elif 'IFCPROPERTYSET' in line_content[1]:
                ifcproduct = 'IFCPROPERTYSET'
                text = line_content[1].replace('IFCPROPERTYSET', '')
                text = text.replace(' (', '(')
                text = text[1:-1]
                text = text.replace('(', '')
                text = text.replace(')', '')
                text = text.replace('\'', '')
                text = text.split(',')
                ifcpara = [text[0],text[1],text[2],text[3],text[4:len(text)-1]]
                list_propertyset.append(IfcDef.IFCPROPERTYSET(line_content[0], ifcproduct, ifcpara))
            elif 'IFCRELDEFINESBYPROPERTIES' in line_content[1]:
                ifcproduct = 'IFCRELDEFINESBYPROPERTIES'
                text = line_content[1].replace('IFCRELDEFINESBYPROPERTIES','')
                text = text.replace(' (', '(')
                text = text[1:-1]
                text = text.replace('(', '')
                text = text.replace(')', '')
                text = text.split(',')
                ifcpara = [text[0],text[1],text[2],text[3],text[4:len(text)-1],text[len(text)-1]]
                list_reldefinesbyproperties.append(IfcDef.IFCRELDEFINESBYPROPERTIES(line_content[0], ifcproduct, ifcpara))
            elif 'IFCPROPERTYSINGLEVALUE' in line_content[1]:
                ifcproduct = 'IFCPROPERTYSINGLEVALUE'
                text = line_content[1].replace('IFCPROPERTYSINGLEVALUE', '')
                text = text.replace(' (', '(')
                text = text.replace('(', '')
                text = text.replace(')', '')
                text = text.replace('IFCLABEL', '')
                text = text.replace('\'', '')
                text = text.split(',')
                ifcpara = [text[0], text[1], text[2], text[3]]
                list_propertysinglevalue.append(IfcDef.IFCPROPERTYSINGLEVALUE(line_content[0], ifcproduct, ifcpara))
        if content_typ == 'end_data':
            print('IFC Datei vollständig eingelesen')
            break
    IFC = IFCFILE(list_element,list_propertyset,list_reldefinesbyproperties,list_propertysinglevalue)
    print(str(len(list_element))+' Objekte ('+ifc_element_type+') im Modell gefunden')
    number_last = int(ifccontent[(len(ifccontent)-1)][0].replace('#', ''))
    number_new = number_last
    file_in_inhalt.strip()
    file_out.write(file_in_inhalt[0:-26])
    file_out.write("\n")

    # comparison of model and excel sheet
    for sheet_element in excel_sheet:
        number_new, file_out = addIFCProp(number_new, file_out, path, file_excel, sheet_element, IFC)

    file_out.write("ENDSEC;\n")
    file_out.write("END-ISO-10303-21;")
    file_out.close()
