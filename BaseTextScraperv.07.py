import openpyxl as op
import pandas as pd
import xlsxwriter as xw
import glob, re, os


all_nec_files = glob.glob('C:\\Users\\dav\\statistik\\*.txt')

here_is_everything = []
for file in all_nec_files:
    txt = open(file, 'r+')
    here_is_everything.append(txt.read())

what_upQuestionmarks = []
what_up109a = []
what_up109b = []
what_up109j = []
what_upSonstige = []
what_upWeb = []
what_upGesamt = []

for file in here_is_everything:
    what_upQuestionmarks.append(re.search('\?\?\?(.*?)109a', file))
    what_up109a.append(re.search('109a(.*?)109b', str(file)))
    what_up109b.append(re.search('109b(.*?)109h', str(file)))
    what_up109j.append(re.search('109j(.*?)Sonstige', str(file)))
    what_upSonstige.append(re.search('Sonstige(.*?)Web', str(file)))
    what_upWeb.append(re.search('Web(.*?)Gesamt', str(file)))
    what_upGesamt.append(re.search('Gesamt(.*?)_', file))

print(here_is_everything[0])

from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = 'new_technique'

sheet = wb['new_technique']

counter = 1
for element in what_upGesamt:
    ws['A' + str(counter)] = str(element) # assign every {counter} cell a value
    counter += 20

wb.save('new_approach.xlsx')
