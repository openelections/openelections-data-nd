import xlrd
import csv

office = 'State House'
district = 47
filename = 'staterep%s_d_precinct.csv' % str(district)
party = 'DEM'

headers = ['county', 'precinct', 'office', 'district', 'party', 'candidate', 'votes']
book = xlrd.open_workbook("Legislative District Precinct Results.xlsx")
total_sheets = book.nsheets
with open(filename, 'wt') as csvfile:
    w = csv.writer(csvfile)
    w.writerow(headers)
    for sheet in range(0, total_sheets):
        sh = book.sheet_by_index(sheet)
        header = sh.row(6)
        candidates = [x.value for x in header[2:]]
        for r in range(7, sh.nrows):
            row = sh.row(r)
            county = sh.name
            precinct = row[1].value
            if precinct == 'TOTALS':
                continue
            cand_votes = zip(candidates, [x.value for x in row[2:]])
            for cand in cand_votes:
                w.writerow([county, precinct, office, district, party, cand[0].strip(), int(cand[1])])
