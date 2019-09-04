# -*- coding: utf-8 -*-
"""
Created on Sun Sep  1 00:06:22 2019

@author: John
"""

import openpyxl, pprint
print('Opening workbook......')
wb = openpyxl.load_workbook(r'C:\Users\John\Desktop\Automate_the_Boring_Stuff_onlinematerials_v.2\automate_online-materials\censuspopdata.xlsx')
sheet=wb['Population by Census Tract']
countydata = {}

# TODO fill in countydata with each countys population and tracts
print('Reading rows.....')
for row in range(2, sheet.max_row +1):
    state = sheet['B'+str(row)].value
    county = sheet['C'+str(row)].value
    pop = sheet['D'+str(row)].value
    
    countydata.setdefault(state,{})
    countydata[state].setdefault(county,{'tracts':0, 'pop':0})
    countydata[state][county]['tracts'] +=1
    countydata[state][county]['pop'] += int(pop)

print('Writing results........')
resultfile = open('census2010.txt','w')
resultfile.write('alldata = \n'+ pprint.pformat(countydata))
resultfile.close()
print('Done')

    