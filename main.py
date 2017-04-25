import xlrd
import shutil
import re

def case_finder(str_list=[]):
    productNum = len(str_list)
    if productNum == 3:
        pattern = re.compile('duanjiamin(.*)?-after')
        match = pattern.match(str_list[2])
        if match:
            return 1
    elif productNum == 4:
        pattern = re.compile('duanjiamin(.*)?-quick-on-after')
        match = pattern.match(str_list[3])
        if match:
            return 2
    elif productNum == 2:
        pattern=re.compile('duanjiamin(.*)?-unnormal-2')
        match=pattern.match(str_list[1])
        if match :
            return 3
    else :
        return 0



def main():
    scr = xlrd.open_workbook('714.xlsx')
    table = scr.sheet_by_name('TQLSIJ')
    path="F:\\xlrd_test\\714\\"
    new_path="F:\\xlrd_test\\147\\"
    rows=table.nrows
    for i in range(rows):
        cell_value=table.cell(i,9).value
        str_list=cell_value.split()
        case_num=case_finder(str_list)
        if case_num == 0:
            pass
        else:
            for djm in range(len(str_list)):
                shutil.copyfile(path+str_list[djm]+'.bmp',new_path+str(i+1)+'duanjiamin_'+str(djm+1)+'.bmp')


if __name__ =="__main__":
    main()


