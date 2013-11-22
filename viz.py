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
        if '+BANI' in s.name:
            sheets[s.name[:-5]] = {}
            for row in range(3, s.nrows - 1):
                if s.cell(row, 1).value == 'TOTAL':
                    break
                sheets[s.name[:-5]][s.cell(row, 1).value] = {'cant': s.cell(row, 6).value,
                                                        'bani': s.cell(row, 7).value}
    return sheets


def reduce_months(acc, new_month):
    for person in new_month:
        if person not in acc:
            acc[person] = {}
        for place in new_month[person]:
            if place not in acc[person]:
                acc[person][place] = []
            acc[person][place].append(new_month[person][place])
    return acc

septembrie = read_excel('Gabi septembrie 2013.xls')
octombrie = read_excel('GABI OCTOMBRIE 2013.xls')

combined = reduce(reduce_months, [septembrie, octombrie, septembrie], {})
# print(combined['OLIVIU'])
for place in combined['OLIVIU']:
    plt.plot(map(lambda x: x['cant'], combined['OLIVIU'][place]), label=place)
plt.legend(loc=1)
plt.show()

