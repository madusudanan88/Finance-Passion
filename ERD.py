import xlrd
import datetime
from datetime import date
class erd:
    def __init__(self,file):
        self.file=file
    def get_total_current_value(self):
        Sum=0
        for  i in self.get_category():
            Sum+=self.get_current_value_by_category(i)
        return(Sum)
    def get_total_installment_value(self):
        Sum=0
        for  i in self.get_category():
            Sum+=self.get_installement_value_by_category(i)
        return(Sum)
    def get_category(self):
        wb = xlrd.open_workbook(self.file)
        sheet=wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if(sheet.cell(0,i).value=='Category'):
                j=i
                break
        C=[]
        for i in range(1,sheet.nrows):
            k=sheet.cell(i,j).value
            if(k not in C):
                C.append(k)
        return(C)

    def get_current_value_by_category(self,cat):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Category"):
                k=i
                break
        Sum=0
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,k).value!=cat):
                continue
            Sum+=self.get_current_value_by_acct(sheet.cell(i,j).value)
        return(Sum)


    def get_installement_value_by_category(self,cat):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Amount"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Category"):
                k=i
                break
        Sum=0
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,k).value!=cat):
                continue
            Sum+=int(sheet.cell(i,j).value)
        return(Sum)

    def get_account_list_by_category(self,cat):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Category"):
                k=i
                break
        A=[]
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,k).value!=cat):
                continue
            A.append(sheet.cell(i,j).value)
        return(A)

    def get_amount_by_account(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Amount"):
                k=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            return(sheet.cell(i,k).value)

    def get_maturity_amount_by_account(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Maturity Amount"):
                k=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            return(sheet.cell(i, k).value)

    def get_maturity_date_by_account(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Maturity Date"):
                k=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            s0=xlrd.xldate_as_datetime(sheet.cell(i,k).value,0)
            s1=str(s0.day)+"-"+str(s0.month)+"-"+str(s0.year)
            return(s1)

    def get_start_date_by_account(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Start Date"):
                k=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            s0=xlrd.xldate_as_datetime(sheet.cell(i,k).value,0)
            s1=str(s0.day)+"-"+str(s0.month)+"-"+str(s0.year)
            return(s1)

    def get_remaining_by_account(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Maturity Date"):
                k=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            matdate = xlrd.xldate_as_datetime(sheet.cell(i,k).value,0)
            curdate=[int(i) for  i in date.today().strftime('%d-%m-%Y').split('-')]
            if(curdate[2]==matdate.year):
                m = matdate.month-curdate[1]-1
                if (curdate[0] < matdate.day):
                    m+=1
                return(m)
            elif(curdate[2]<matdate.year):
                ydiff=matdate.year-curdate[2]-1
                mcount=ydiff*12
                m=matdate.month-1
                mcount+=m
                m=12-curdate[1]
                mcount+=m
                if(curdate[0]<matdate.day):
                    mcount+=1
                return(mcount)
            elif(curdate[2]>matdate.year):
                return(-1)

    def get_current_value_by_acct(self,acct):
        wb = xlrd.open_workbook(self.file)
        sheet = wb.sheet_by_name('E-RD')
        j=-1
        m=0
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Account No"):
                j=i
                break
        k=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Start Date"):
                k=i
                break
        l=-1
        for i in range(sheet.ncols):
            if (sheet.cell(0, i).value == "Maturity Date"):
                l=i
                break
        for i in range(1,sheet.nrows):
            if(sheet.cell(i,j).value!=acct):
                continue
            startdate = xlrd.xldate_as_datetime(sheet.cell(i,k).value,0)
            matdate = xlrd.xldate_as_datetime(sheet.cell(i, l).value, 0)
            m4=self.get_amount_by_account(acct)
            if(matdate.year>startdate.year):
                m1=(matdate.year-startdate.year-1)*12
                m1+=(matdate.month-1)
                m1+=(13-startdate.month)
                m2=self.get_remaining_by_account(acct)
                m3=m1-m2
                return(m3*m4)
            else:
                return(0)


