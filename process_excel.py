import xlrd, xlwt, os

datas = []
biaotou = []
path = os.path.abspath('.')
original_file_path = os.path.join(path, 'to_be_processed')
files = os.listdir(original_file_path)
for file in files:
	print('processing ' + file)
	workbook = xlrd.open_workbook(os.path.join(original_file_path, file))
	sheet = workbook.sheet_by_index(0)
	n = sheet.nrows
	biaotou = sheet.row_values(0)
	for i in range(1, n):
		datas.append(sheet.row_values(i))

new_workbook = xlwt.Workbook()
new_sheet = new_workbook.add_sheet('sheet_1')

for i in range(len(biaotou)):
	new_sheet.write(0, i, biaotou[i])

for i in range(0,len(datas)):
	for j in range(len(datas[i])):
		new_sheet.write(i, j, datas[i][j])
new_workbook.save('result.xls')
input('请按回车键结束：')