from openpyxl import load_workbook
from functools import reduce
from collections import defaultdict
import argparse
import os

parser = argparse.ArgumentParser()
parser.add_argument('--separate', action='store_true')

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

def sample_num(x):
	if x <= 15:
		return 2
	elif x <= 25:
		return 3
	elif x <= 90:
		return 5
	elif x <= 150:
		return 8
	elif x <= 280:
		return 13
	elif x <= 500:
		return 20
	elif x <= 1200:
		return 32
	else:
		return 50

def uids_to_str(uids):
	d = defaultdict(list)
	for uid in uids:
		d[uid[:3]].append(int(uid[3:]))
	ret = ''
	first_k = True
	for k, l in d.items():
		x = sorted(l)
		cur = 0
		r = 0
		if first_k:
			ret += k
		else:
			ret += '；' + k
		first_k = False
		first_n = True
		while cur < len(x):
			r = cur
			while r + 1 < len(x) and x[r + 1] <= x[r] + 1:
				r += 1
			low = x[cur]
			high = x[r]
			to_append = '' if first_n else '、'
			if low == high:
				to_append += '{}'.format(low)
			else:
				to_append += '{}~{}'.format(low, high)
			first_n = False
			ret += to_append
			cur = r + 1
	return ret

def date_to_chinese(s):
	l = s.split('.')
	return l[0] + '年' + l[1] + '月' + l[2] + '日'

def process(data, folder, prefix):
	os.makedirs(folder, exist_ok=True)
	max_page = 0
	d = bin_by_date_and_id(data)
	sheet_no = 0
	for date, v in d.items():
		template_wb = load_workbook('a.xlsx')
		sheet = template_wb['施工记录']
		sheet_no += 1
		sheet_no_str = '{:03d}'.format(sheet_no)
		template_wb['报审表']['L5'] = sheet_no_str
		template_wb['检验批']['BF4'] = sheet_no_str
		template_wb['隐蔽工程']['AG3'] = sheet_no_str
		template_wb['报审表']['K21'] = date_to_chinese(date)
		template_wb['检验批']['AY26'] = date_to_chinese(date)
		template_wb['隐蔽工程']['Y5'] = date_to_chinese(date)
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
				sheet.cell(row=get_row(page, offset), column=3).value = row[2].value.replace('.', '/')
				uids.append(row[3].value)
				offset += 1
				if offset >= 20:
					offset -= 20
					page += 1
			if offset != 0:
				offset = 0
				page += 1
		sample_n = sample_num(first_col)
		template_wb['检验批']['BB8'] = first_col
		for i in range(12, 23):
			template_wb['检验批']['AD{}'.format(i)] = sample_n
			template_wb['检验批']['AI{}'.format(i)] = sample_n
			template_wb['检验批']['AL{}'.format(i)] = '检查{}处，合格{}处'.format(sample_n, sample_n)
		sheet.print_area = 'A1:L{}'.format(last_page * 24 + 24)
		sheet0 = template_wb['报审表']
		uid_str = uids_to_str(uids)
		sheet0['C8'] = uid_str
		template_wb['隐蔽工程']['G7'] = uid_str
		max_page = max(max_page, last_page)
		template_wb.save('{}/{}.xlsx'.format(folder, prefix + date[5:]))
	print('Need {} pages.'.format(max_page + 1))

args = parser.parse_args()
if not args.separate:
	data = reduce(lambda x, y: x + y, [list(data_wb[name].rows)[4:] for name in ids])
	data = filter(lambda x: x[0].value is not None, data)
	process(data, 'result', '')
else:
	for name in ids:
		data = list(data_wb[name].rows)[4:]
		data = filter(lambda x: x[0].value is not None, data)
		process(data, 'result/data' + name, name + '单轴检验批')

