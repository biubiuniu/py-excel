# -*- coding: utf-8 -*- 
import xlrd
import xlwt
import glob
import os
import sys
reload(sys)
sys.setdefaultencoding( "utf-8" )
old_filename_list = []
name_list= []
path = 'E:\\test'
count = 0
p = 0
sp_1 = []
def filename():
	global old_filename_list,path,count
	for root, dirs, files in os.walk(path): 
		fileLength = len(files)  #读取目录下的文件数
		if fileLength != 0:
			count = count + fileLength
		for i in range(len(files)):  #获取目录下文件的路径+文件名
			old_filename = str(root) + '/' + files[i]
			old_filename = old_filename.replace('\\','/')
			old_filename_list.append(old_filename)
	print old_filename_list
	# print count
	print "The number of files under <%s> is: %d" %(path,count)

def excelwrite():
	global name_list,p,count,old_filename_list,sp_1
	for file in old_filename_list:
		f = open (file)
		filename_left = f.name.split("/")[-1].split(".xls")[0] #提取文加名，除去后缀名
		print filename_left
		if ( p < count ):
			if ('xls' in file):
				workbook = xlrd.open_workbook(file) #读取文件
				# print workbook.sheet_names()
				sheet2 = workbook.sheet_by_index(1) #读取工作表
				nrows = sheet2.nrows
				# print nrows
				book = xlwt.Workbook(encoding='utf-8',style_compression=0)
				sheet = book.add_sheet('1',cell_overwrite_ok=True)
				list = [3,5,7,11,12]
				list_1 = [1,2,3,4,5]
				s = 0

				# sheet2 = workbook.sheets()[0]
				sp_1 = []
				for sp in range(nrows-1):
					danyuan = sheet2.cell(sp,5).value.encode('utf-8')
					if ( danyuan != "[低]" ):
						sp_1.append(sp)
					sp = sp +1
				print sp_1

				for i in list: #读取列数据
					cols = sheet2.col_values(i) #读取列数据
					# m = 1
					# for y in cols[3:]: #从第四列开始读取写入
					# 	if m in sp_1:
					# 		a = str(y).encode("utf-8") 
					# 		b = a + '\n'
					# 		name_list.append(b) #取得数据list
					# 		print b
					# 	m = m + 1
					# print y
					for m in sp_1[1:]:
							a = sheet2.cell(m,i).value.encode('utf-8')
							name_list.append(a)
					for x in range(len(name_list)): #将上面读取的数据循环写入
						sheet.write(x,list_1[s],name_list[x]) #行数为X 列数为i 插入数据为name_list[x]
						sheet.write(x,0,filename_left)
					name_list = []
					s = s + 1
			book.save('E:\\test\\'+str(p)+'.xls')
			p = p + 1

if __name__ == '__main__':
	filename()
	excelwrite()