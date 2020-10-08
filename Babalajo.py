from  Tkinter import *

import xlrd
import xlwt
from xlwt import Workbook
from xlwt import Workbook,Formula

wb=Workbook()
sheet1=wb.add_sheet('Moneyrecords')
sheet1.write(0,0,'DATE')
sheet1.write(0,1,'NAMES')
sheet1.write(0,2,'HOME ADDRESS')
sheet1.write(0,3,'SEX')
sheet1.write(0,4,'PHONE NOS')
sheet1.write(0,5,'BANK DETAILS')


sheet1.write(0,6,'AMOUNT-DAY1')
sheet1.write(103,5,"TOTAL PROFIT")
sheet1.write(0,37,"SUM")
sheet1.write(0,7,'DAY2')
sheet1.write(0,8,'DAY3')
sheet1.write(0,9,'DAY4')                                 
sheet1.write(0,10,'DAY5')
sheet1.write(0,11,'DAY6')
sheet1.write(0,12,'DAY7')
sheet1.write(0,13,'DAY8')
sheet1.write(0,14,'DAY9')
sheet1.write(0,15,'DAY10')
sheet1.write(0,16,'DAY11')
sheet1.write(0,17,'DAY12')
sheet1.write(0,18,'DAY13')
sheet1.write(0,19,'DAY14')
sheet1.write(0,20,'DAY15')
sheet1.write(0,21,'DAY16')
sheet1.write(0,22,'DAY17')
sheet1.write(0,23,'DAY18')
sheet1.write(0,24,'DAY19')
sheet1.write(0,25,'DAY20')
sheet1.write(0,26,'DAY21')
sheet1.write(0,27,'DAY21')
sheet1.write(0,28,'DAY22')
sheet1.write(0,29,'DAY23')
sheet1.write(0,30,'DAY24')
sheet1.write(0,31,'DAY25')
sheet1.write(0,32,'DAY26')
sheet1.write(0,33,'DAY27')
sheet1.write(0,34,'DAY28')
sheet1.write(0,35,'DAY29')
sheet1.write(0,36,'DAY30')
                            
sheet1.col(0).width=3000
sheet1.col(1).width=6000
sheet1.col(2).width=4000
  
sheet1.col(3).width=4000
sheet1.col(4).width=7000
sheet1.col(5).width=8000
sheet1.col(6).width=7000


class Application(Frame):

    def __init__(self,master):
        Frame.__init__(self,master)
        self.grid()
        self.create_widgets()
     
    
        

    def create_widgets(self):
        self.instruction1=Label(self,text="-----THRIFT SAVER CALCULATOR-----(BABAA ALAJOO)-----",bg="powder blue",font=("bold")).grid(row = 0,column =0,columnspan =9,  sticky = W)
       
        self.instruction=Label(self,text="ENTER THE CUSTOMER NUMBER:",bg="powder blue",fg="brown").grid(row = 2,column =0,columnspan =7,  sticky = W)
        self.display1=IntVar()
        self.num=Entry(self,textvariable=self.display1,bd=6,bg="powder blue").grid(row=2, column=7,columnspan=12, sticky =W)
        
        self.instruction2=Label(self,text="ENTER THE DATE THE CUSTOMER PAID:",bg="powder blue",fg="brown").grid(row = 4,column =0,columnspan =7,  sticky = W)
        self.display2=StringVar()
        self.date=Entry(self,textvariable=self.display2 ,bd=10,bg="powder blue").grid(row=4, column=7,columnspan=7, sticky =W)
        
        
        self.instruction3=Label(self,text="NAME OF THE CUSTOMER:",fg="brown",bg="powder blue").grid(row = 6,column = 0,columnspan =7,  sticky = W)
        self.display3=StringVar()
        self.name=Entry(self,textvariable=self.display3,bd=10,bg="powder blue").grid(row =6, column =7,columnspan=12, sticky = W)
        
        self.instruction4=Label(self,text="ADDRESSS OF THE CUSTOMER:",fg="brown",bg="powder blue").grid(row=7,column =0,columnspan = 7,sticky =W)
        self.display4=StringVar()
        self.address=Entry(self,textvariable=self.display4,bd=10,bg="powder blue").grid(row=7,column = 7,columnspan = 12,sticky= W)
        
        
        self.instruction5=Label(self,text="AMOUNT PAID BY THE CUSTOMER:",fg="brown",bg="powder blue").grid(row=9,column =0,columnspan = 7,sticky =W)
        self.display5=IntVar()
        self.amount=Entry(self,textvariable=self.display5,bd=10,bg="powder blue").grid(row=9,column = 7,columnspan = 12,sticky= W)

        self.instruction4=Label(self,text="TOTAL PREVIOUS AMOUNT:",fg="brown",bg="powder blue").grid(row=20,column =0,columnspan = 7,sticky =W)
        self.display8=IntVar()
        self.costumernos=Entry(self,textvariable=self.display8,bd=10,bg="powder blue").grid(row=20,column = 8,columnspan = 12,sticky= W)

        self.instruction10=Label(self,text="ENTER THE DAYNUMBER:",bg="powder blue",fg="brown").grid(row=8,column =0,columnspan = 7,sticky =W)
        
        self.display10=IntVar()
        self.costumernos=Entry(self,textvariable=self.display10,bd=6,bg="powder blue",).grid(row=8,column = 7,columnspan = 12,sticky= W)

        Label(self,text="GENDER",fg="brown",bg="powder blue").grid(row =10,column =0,sticky=W)
        self.favourite=IntVar()
        
        
        Radiobutton(self,text="Male",variable=self.favourite,command = self.update_sex,padx = 20,value=1).grid(row=10,column=5,sticky= W)
        Radiobutton(self,text="Female",variable=self.favourite,command = self.update_sex,padx = 20,value=2).grid(row=10,column=9,sticky= W)

        self.instruction5=Label(self,text="PHONE NOS OF THE CUSTOMER:",fg="brown",bg="powder blue").grid(row=12,column =0,columnspan = 7,sticky =W)
        self.display31=IntVar()
        self.amount=Entry(self,textvariable=self.display31,bd=10,bg="powder blue").grid(row=12,column = 7,columnspan = 12,sticky= W)
        
        self.instruction5=Label(self,text="BANK ACCOUNT DETAILS FOR CUSTOMER:",fg="brown",bg="powder blue").grid(row=13,column =0,columnspan = 7,sticky =W)
        self.display30=StringVar()
        self.amount=Entry(self,textvariable=self.display30,bd=10,bg="powder blue").grid(row=13,column = 7,columnspan = 12,sticky= W)
        
        
        self.button1 = Button(self,text="ADD CUSTOMER DATA",bd=5,fg="yellow",bg="navy blue",command = self.update_data1).grid(row=19,column=0,sticky=E)
        self.button1 = Button(self,text="SEND DATA TO DATABASE",bd=10,fg="white",bg="red",command = self.update_data2).grid(row=25,column=0,sticky=E)
        self.button1 = Button(self,text="ADD AMOUNT FOR CUSTOMER",fg="black",bg="green",bd="5",command = self.update_data3).grid(row=18,column=0,sticky=E)
        
        self.instruction9=Label(self,text="TOTAL COMMISSION FOR BABA ALAJO:",fg="brown",bg="powder blue").grid(row=22,column =0,columnspan = 7,sticky =W)
        
        self.text1=Text(self,width=20,height=1,wrap=WORD,bd="2")
        self.text1.grid(row=22,column=5,sticky=W)
        

    def update_data1(self):
        i=self.display1.get()
        date=self.display2.get()
        sheet1.write(i,0,date)
        name=self.display3.get()
        sheet1.write(i,1,name)
        address=self.display4.get()
        sheet1.write(i,2,address)
        email=self.display31.get()
        sheet1.write(i,4,email)
        bankdetails=self.display30.get()
        sheet1.write(i,5,bankdetails)
        
        
        
        amount=int(self.display5.get())
        sum1 = self.display8.get()
        sum1 = sum1 + amount
        self.text1.delete(0.0,END)
        self.text1.insert(0.0, sum1)      
        wb.save('RECORD DATABASE.xls')
    
   
                    
           
    def update_sex(self):
        j=self.display1.get()
        if self.favourite.get()==1:
            sex="M"
            sheet1.write(j,3,sex)
        else:
            sex="F"
            sheet1.write(j,3,sex)
        wb.save('RECORD DATABASE.xls')
        
    
                  
    def update_data2(self):
                     
       
        sheet1.write(103,6,Formula('sum(G2:G102)'))
        sheet1.write(1,37,Formula('sum(H2:AK2)'))
        sheet1.write(2,37,Formula('sum(H3:AK3)'))
        sheet1.write(3,37,Formula('sum(H4:AK4)'))
        sheet1.write(4,37,Formula('sum(H5:AK5)'))
        sheet1.write(5,37,Formula('sum(H6:AK6)'))
        sheet1.write(6,37,Formula('sum(H7:AK7)'))
        sheet1.write(7,37,Formula('sum(H8:AK8)'))
        sheet1.write(8,37,Formula('sum(H9:AK9)'))
        sheet1.write(9,37,Formula('sum(H10:AK10)'))
        sheet1.write(10,37,Formula('sum(H11:AK11)'))
        sheet1.write(11,37,Formula('sum(H12:AK12)'))
        sheet1.write(12,37,Formula('sum(H13:AK13)'))
        sheet1.write(13,37,Formula('sum(H14:AK14)'))
        sheet1.write(14,37,Formula('sum(H15:AK15)'))
        sheet1.write(15,37,Formula('sum(H16:AK16)'))
        sheet1.write(16,37,Formula('sum(H17:AK17)'))
        sheet1.write(17,37,Formula('sum(H18:AK18)'))
        sheet1.write(18,37,Formula('sum(H19:AK19)'))
        sheet1.write(19,37,Formula('sum(H20:AK20)'))
        sheet1.write(20,37,Formula('sum(H21:AK21)'))
        sheet1.write(21,37,Formula('sum(H22:AK22)'))
        sheet1.write(22,37,Formula('sum(H23:AK23)'))
        sheet1.write(23,37,Formula('sum(H24:AK24)'))
        sheet1.write(24,37,Formula('sum(H25:AK25)'))

        sheet1.write(25,37,Formula('sum(H26:AK26)'))
        sheet1.write(26,37,Formula('sum(H27:AK27)'))
        sheet1.write(27,37,Formula('sum(H28:AK28)'))
        sheet1.write(28,37,Formula('sum(H29:AK29)'))
        sheet1.write(29,37,Formula('sum(H30:AK30)'))
        sheet1.write(30,37,Formula('sum(H31:AK31)'))
                                                     
        
         
        wb.save('RECORD DATABASE.xls')

        
    def update_data3(self):
        i=self.display1.get()
        
        j=self.display10.get()
        
        k=j+5
        
        amount=int(self.display5.get())
        
        sheet1.write(i,k,amount)
        
        wb.save('RECORD DATABASE.xls')


       
        
        

        
    
    
        
       
        
        
       
  
       
        
        
                 
root = Tk()
root.title("Baba alajo")
root.geometry("650x550")
root.configure(background='light blue')
app=Application(root)
root.mainloop()        
