import xlwt
from xlrd import open_workbook
import time as t
from xlutils.copy import copy
while(1):
      print("...Welcome...")
      book=open_workbook("xclwirte.xls")
      wb=copy(book)
      s=wb.get_sheet(0)
      sheet=book.sheets()
      p=sheet[0].nrows
      name_list=[]
      pswd_list=[]
      balance_list=[]
      print('1. Withdraw','\n2. Balance Enquiry','\n3. Cancel')
      operation=int(input("Choose Operation : "))
      if(operation==1):
            pin=int(input("Enter your pin here : "))
            for i in (sheet):
                  for j in range(p):
                        name=i.cell(j,0).value
                        name_list.append(name)
                        pswd=i.cell(j,1).value
                        pswd_list.append(pswd)
                        balance=i.cell(j,2).value
                        balance_list.append(balance)
            for n in range(j+1):
                  b=pswd_list[n]
                  c=name_list[n]
                  if(pin==pswd_list[n]):
                        print("Welcome ",name_list[n])
                        amount=int(input("Enter your amount to withdraw : "))
                        balance_list[n]=balance_list[n]-amount
                        s.write(n,2,balance_list[n])
                        print(balance_list[n],"\n")
            wb.save("xclwirte.xls")
      elif(operation==2):
            pin=int(input("Enter your pin here :="))
            for i in (sheet):
                  for j in range(p):
                        name=i.cell(j,0).value
                        name_list.append(name)
                        pswd=i.cell(j,1).value
                        pswd_list.append(pswd)
                        balance=i.cell(j,2).value
                        balance_list.append(balance)
            for n in range(j+1):
                  b=pswd_list[n]
                  c=name_list[n]
                  if(pin==pswd_list[n]):
                        print("Welcome ",name_list[n])
                        print("Your current balance is :=",balance_list[n],"Rs.\n")
      else:
            print("Thank You for choosing us :)")
            break
            
            
           
