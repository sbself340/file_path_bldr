# -*- coding: utf-8 -*-
"""
Created on Wed Apr 11 09:51:13 2018

@author: BSelf
"""

import pandas as pd
import os

def locations():
    parent = str(input("Address to Parent Folder: "))
    destination = str(input("Address to Destination Folder"
                            "(If same as Parent just hit enter): "))
    
#    parent = r'C:\Users\BSelf\Desktop\EQT Scripts'
#    destination = r'C:\Users\BSelf\Desktop\EQT Scripts'
    return parent, destination

def find_paths(location):
    print('Finding Paths')
    paths = []
    for root, dirs, files in os.walk(location):
        paths.append(root)
    return paths

def list_files(paths):
    print('Finding Files')
    name = []
    fullname = []
    loc = []
    for location in paths:
        os.chdir(location)
        doc_list = os.listdir()
        for doc in doc_list:
            if '.' in doc[-5:]:
                name.append(doc)
                fullname.append(str(location + '\\' + doc))
                loc.append(location[3:])
    return name, fullname, loc

def dashes(loc):
    dashnum = []
    for locations in loc:
        dashnum.append(locations.count('\\'))
    return dashnum

def path_parser(loc, fullname, name):
    dash_num = dashes(loc)
    df = pd.DataFrame({'File_Name': name, 'Full_Path': fullname})
    locburn = list(loc)
    number1 = max(dash_num) + 1
    for i in range(number1):
        lis = []
        lisname = 'Path_' + str(i)
        number2 = len(dash_num)
        for j in range(number2):
            if '\\' in locburn[j]:
                count = locburn[j].find('\\')
                lis.append(locburn[j][:count])
                count += 1
                locburn[j] = locburn[j][count:]
            elif locburn[j] != '':
                lis.append(locburn[j])
                locburn[j] = ''
            else:
                lis.append('')
        lis = pd.Series(lis)
        df[lisname] = lis
    return df

def xlsx_writer(df, destination):
    writer = pd.ExcelWriter('Indexed_Files.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()

def main():
    parent, destination = locations()
#    print('{}\n{}'.format(parent, destination))
#    os.chdir(parent)
    paths = find_paths(parent)
    name, fullname, loc = list_files(paths)
#    dash_num = dashes(loc)
    df = path_parser(loc, fullname, name)
    xlsx_writer(df, destination)
    print('File Created')
    
    

if __name__ == '__main__':
    main()