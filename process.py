from openpyxl import load_workbook
from functools import reduce
from collections import defaultdict

data_wb = load_workbook('b.xlsx', data_only=True)
ids = data_wb.sheetnames

def bin_by_date_and_id(data):
	ret = defaultdict(lambda: defaultdict(list))
	for row in data:
		machine_id = row[1].value
		date = row[2].value
		ret[date][machine_id].append(row)
	return ret

def get_row(page, offset):
	return page * 24 + offset + 5

def uids_to_str(uids):
	d = defaultdict(list)
	for uid in uids:
		d[uid[:3]].append(int(uid[3:]))
	ret = ''
	first_append = True
	for k, l in d.items():
		x = sorted(l)
		cur = 0
		r = 0
		while cur < len(x):
			r = cur
			while r + 1 < len(x) and x[r + 1] <= x[r] + 1:
				r += 1
			low = x[cur]
			high = x[r]
			to_append = '' if first_append else ', '
			if low == high:
				to_append += '{}{}'.format(k, low)
			else:
				to_append += '{}{}~{}{}'.format(k, low, k, high)
			first_append = False
			ret += to_append
			cur = r + 1
	return ret

def process(data):
	max_page = 0
	d = bin_by_date_and_id(data)
	for date, v in d.items():
		template_wb = load_workbook('a.xlsx')
		sheet = template_wb['施工记录']
		page = 0
		offset = 0
		first_col = 0
		last_page = 0
		uids = []
		for rows in v.values():
			for row in rows:
				first_col += 1
				last_page = page
				sheet.cell(row=get_row(page, offset), column=1).value = first_col
				for i in range(1, 12):
					sheet.cell(row=get_row(page, offset), column=i+1).value = row[i].value
				uids.append(row[3].value)
				offset += 1
				if offset >= 20:
					offset -= 20
					page += 1
			if offset != 0:
				offset = 0
				page += 1
		sheet.print_area = 'A1:L{}'.format(last_page * 24 + 24)
		sheet0 = template_wb['报审表']
		sheet0['B8'] = '我方已完成   {}   单轴水泥搅拌桩的施工工作，经自检合格，请予以审查或验收。'.format(uids_to_str(uids))
		max_page = max(max_page, last_page)
		template_wb.save('result/{}.xlsx'.format(date))
	print('Need {} pages.'.format(max_page + 1))

data = reduce(lambda x, y: x + y, [list(data_wb[name].rows)[4:] for name in ids])
data = filter(lambda x: x[0].value is not None, data)
process(data)

