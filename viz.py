#!/usr/bin/python
import sys
from xlrd import open_workbook
import xlwt
from pprint import pprint
import numpy as np
import matplotlib.pyplot as plt

try:
    reduce
except NameError:
    from functools import reduce
# from mpltools import style

# style.use(['ggplot', 'pof'])


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
                                                             'bani': s.cell(row, 7).value,
                                                             'rebut': s.cell(row, 8).value}
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

if len(sys.argv) == 1:
    sys.argv += ['Gabi septembrie 2013.xls', 'GABI OCTOMBRIE 2013.xls']

combined = reduce(reduce_months, map(read_excel, sys.argv[1:]), {})
#print(combined['OLIVIU'])
for person in combined:
    print("========== %s ==========" % person)
    for place in combined[person]:
        x = combined[person][place]
        try:
            cant = (x[1]['cant'] - x[0]['cant']) / x[0]['cant'] * 100
            bani = (x[1]['bani'] - x[0]['bani']) / x[0]['bani'] * 100
            medie = x[1]['bani'] / (x[1]['cant'] - x[1]['rebut'])
            print("| {0: ^18.18} | {1: ^8.3f} | {2: ^8.3f} | {3: ^8.3f} |".format(place, cant, bani, medie))
        except:
            pass
            #print("| {0: ^18.18} | {1: ^7.3f} | {2: ^7.3f} |".format(place, x[1]['cant'], x[1]['bani']))
    #    plt.plot(list(map(lambda x: x['cant'], combined[person][place])), label=place)
    #plt.show()
#plt.legend(loc=2).draggable(True)


