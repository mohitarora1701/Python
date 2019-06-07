import xlwt
from xlrd import open_workbook
import time as t
from xlutils.copy import copy
import random as r
while(1):
      print("...Welcome...")
      book=open_workbook("xclwirte.xls")
      wb=copy(book)
      s=wb.get_sheet(0)
      sheet=book.sheets()
      p=sheet[0].nrows
      name_list=[]
      number_list=[]
      balance_list=[]
      print("1. Add money","\n2. Change Password","\n3. New Account","\n4. Exit")
      operation=int(input("Enter your operation : "))
      if(operation==1):
            user_name=input("Enter user name =")
            user_number=int(input("Enter user phone number ="))
            for i in (sheet):
                  for j in range(p):
                        name=i.cell(j,0).value
                        name_list.append(name)
                        balance=i.cell(j,2).value
                        balance_list.append(balance)
                        number=i.cell(j,3).value
                        number_list.append(number)
                  for n in range(j+1):
                        b=number_list[n]
                        c=name_list[n]
                        if(user_name==name_list[n] and user_number==number_list[n]):
                              print("User name =",user_name,"\nUser number =",user_number)
                              add=int(input("Enter amount to add ="))
                              total=add+balance_list[n]
                              s.write(n,2,total)
                              print("Money added \n")
                              print("Now your current balance is =",total)
                  wb.save("xclwirte.xls")
      elif(operation==2):
            user_name=input("Enter user name =")
            user_number=int(input("Enter user phone number ="))
            for i in (sheet):
                  for j in range(p):
                        name=i.cell(j,0).value
                        name_list.append(name)
                        balance=i.cell(j,2).value
                        balance_list.append(balance)
                        number=i.cell(j,3).value
                        number_list.append(number)
            for n in range(j+1):
                  b=number_list[n]
                  c=name_list[n]
                  if(user_name==name_list[n] and user_number==number_list[n]):
                        print("User name =",user_name,"\nUser number =",user_number)
                        pswd=r.randint(999,10000)
                        s.write(n,1,pswd)
                        print("Your password is set \nYour new password is =",pswd,"\n")
            wb.save("xclwirte.xls")
      elif(operation==3):
            print("Please fill below details..")
            new_name=input("Enter name = ")
            new_number=int(input("Enter phone number = "))
            new_pswd=r.randint(999,10000)
            new_amount=int(input("Enter amount to deposite = "))
            s.write(p,0,new_name)
            s.write(p,1,new_pswd)
            s.write(p,2,new_amount)
            s.write(p,3,new_number)
            print("Account creating....")
            for f in range(6):
                  print("..")
                  t.sleep(0.5)
            print("Name = {} \nNumber = {} \nPassword = {} \nBalance = {}".format(new_name,new_number,new_pswd,new_amount))
            print("You want to create ?","\nY. Yes \nN. No")
            ans=input("=>")
            if(ans=="Y"):
                  wb.save("xclwirte.xls")
                  print("Congratulation and Thank you for join us :) \n")
            elif(ans=="N"):
                  pass
                  print("\n")
            
      elif(operation==4):
            print("Are you sure you want to exit ?")
            print("Y. Yes \nN. No")
            answer=input("=>")
            if(answer=="Y"):
                  print("Successfully Exit")
                  break
            elif(answer=="N"):
                  pass
                  print("\n")
      else:
            print("Please select valid operation")
            pass
            print("\n")
