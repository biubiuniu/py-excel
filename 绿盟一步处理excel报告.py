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
		fileLength = len(files)  #��ȡĿ¼�µ��ļ���
		if fileLength != 0:
			count = count + fileLength
		for i in range(len(files)):  #��ȡĿ¼���ļ���·��+�ļ���
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
		filename_left = f.name.split("/")[-1].split(".xls")[0] #��ȡ�ļ�������ȥ��׺��
		print filename_left
		if ( p < count ):
			if ('xls' in file):
				workbook = xlrd.open_workbook(file) #��ȡ�ļ�
				# print workbook.sheet_names()
				sheet2 = workbook.sheet_by_index(1) #��ȡ������
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
					if ( danyuan != "[��]" ):
						sp_1.append(sp)
					sp = sp +1
				print sp_1

				for i in list: #��ȡ������
					cols = sheet2.col_values(i) #��ȡ������
					# m = 1
					# for y in cols[3:]: #�ӵ����п�ʼ��ȡд��
					# 	if m in sp_1:
					# 		a = str(y).encode("utf-8") 
					# 		b = a + '\n'
					# 		name_list.append(b) #ȡ������list
					# 		print b
					# 	m = m + 1
					# print y
					for m in sp_1[1:]:
							a = sheet2.cell(m,i).value.encode('utf-8')
							name_list.append(a)
					for x in range(len(name_list)): #�������ȡ������ѭ��д��
						sheet.write(x,list_1[s],name_list[x]) #����ΪX ����Ϊi ��������Ϊname_list[x]
						sheet.write(x,0,filename_left)
					name_list = []
					s = s + 1
			book.save('E:\\test\\'+str(p)+'.xls')
			p = p + 1

if __name__ == '__main__':
	filename()
	excelwrite()