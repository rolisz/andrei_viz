#!/usr/bin/python
from xlrd import open_workbook
from pprint import pprint
import numpy as np
import matplotlib.pyplot as plt

from mpltools import style

style.use(['ggplot', 'pof'])


def read_excel(f):
    wb = open_workbook(f, ragged_rows=True)
    sheets = {}
    for s in wb.sheets():
        if '+BANI' not in s.name:
            sheets[s.name] = {}
            for row in range(3, s.nrows - 1):
                sheets[s.name][s.cell(row, 1).value] = s.cell(row, 6).value
    return sheets

septembrie = read_excel('Gabi septembrie 2013.xls')
octombrie = read_excel('GABI OCTOMBRIE 2013.xls')

combined = {}

for person in set(septembrie.iterkeys()).union(octombrie.iterkeys()):
    combined[person] = {}
    places = set()
    for month in [septembrie, octombrie]:
        if person in month:
            places = places.union(month[person].iterkeys())
    for place in places:
        diff = []
        for month in [septembrie, octombrie]:
            if place in month[person]:
                diff.append(month[person][place])
            else:
                diff.append(0)
        combined[person][place] = diff

print(combined['OLIVIU'])
for place in combined['OLIVIU']:
    plt.plot(combined['OLIVIU'][place])
plt.show()

