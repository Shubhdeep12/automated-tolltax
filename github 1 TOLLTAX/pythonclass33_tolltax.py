from tkinter import *
import time as t
'''import RPi.GPIO as rp
rp.setwarnings(False)'''
root=Tk()
root.geometry('600x500')
root.title('TOLL PLAZA NH-14')
f1=Frame(root)
f2=Frame(root)
f3=Frame(root)
f4=Frame(root)
f5=Frame(root)
f6=Frame(root)
f7=Frame(root)
f8=Frame(root)
f9=Frame(root)
f10=Frame(root)
def r_frame(frame):
    frame.tkraise()
    for f in (f1,f2,f3,f4,f5,f6,f7,f8,f9,f10):
        f.grid(row=0,column=0,sticky="news")
def rfid():
    #for rfid card information
    '''ser=0
    import serial
    ser=serial.Serial('/dev/ttyUSB0',9600)'''
    ser=input('enter id')
    import xlrd
    wb=xlrd.open_workbook(r'F:\python\shubh-python\idpass2.xls')
    s=wb.sheet_by_index(0)
    aa=s.col_values(0)
    bb=s.col_values(1)
    cc=s.col_values(2)
    dd=s.col_values(3)
    ee=s.col_values(4)
    ff=s.col_values(5)
    for i in range(0,len(aa)):
        if(aa[i]==ser):
            r_frame(f2)
            Label(f2,text=' ').grid(row=1,column=0)
            Label(f2,text=' ').grid(row=2,column=0)
            Label(f2,text='PLEASE SWIPE YOUR CARD :)',font='Helvetica 30 bold ').grid(row=3,column=1)
            Label(f2,text=' ').grid(row=4,column=0)
            Label(f2,text=' ').grid(row=5,column=0)
            Label(f2,text='Choose your option',font='Helvetica 20 ').grid(row=6,column=0)
            Label(f2,text=' ').grid(row=7,column=0)
            Label(f2,text=' ').grid(row=8,column=0)
            var1 = IntVar()
            Checkbutton(f2, text="ONE WAY",font='Helvetica 10 ', variable=var1).grid(row=9,column=0, sticky=W)
            var2 = IntVar()
            Checkbutton(f2, text="TWO WAY",font='Helvetica 10 ', variable=var2).grid(row=11,column=0, sticky=W)
            Label(f2,text=bb[i]).grid(row=7,column=1)
            Label(f2,text=cc[i]).grid(row=9,column=1)
            Label(f2,text=dd[i]).grid(row=10,column=1)
            Label(f2,text=ee[i]).grid(row=8,column=1)
            Label(f2,text=' ').grid(row=12,column=0)
            Label(f2,text=' ').grid(row=13,column=0)
            Button(f2,text='Submit',font='Helvetica 13',command=lambda:check(i,var1.get(),var2.get())).grid(row=14,column=0)
            break
        
        if(i==2):
            if(aa[i]!=ser):
                r_frame(f10)
                Label(f10,text=' ').grid(row=4,column=0)
                Label(f10,text=' ').grid(row=5,column=0)
                Label(f10,text='INVALID ID\n\n  Please give cash to employee\n\n\n\nENJOY THE RIDE :)',font='Helvetica 20 ').grid(row=6,column=0)
                Label(f10,text=' ').grid(row=7,column=0)
                Label(f10,text=' ').grid(row=8,column=0)
                Button(f10,text='Exit',font='Helvetica 13',command=lambda:r_frame(f1)).grid(row=9,column=0)
r_frame(f1)
Label(f1,text='WELCOME\n\n   TO\n\n  NH-14 TOLL PLAZA',font='Helvetica 30 bold ').grid(row=1,column=1)
Label(f1,text=' ').grid(row=0,column=1)
Label(f1,text=' ').grid(row=1,column=0)
Label(f1,text=' ').grid(row=3,column=0)
Label(f1,text=' ').grid(row=4,column=0)
Label(f1,text=' ').grid(row=5,column=0)
Label(f1,text=' ').grid(row=6,column=0)
Button(f1,text='SWIPE \nCARD',font='Helvetica 15',command=lambda:rfid()).grid(row=7,column=2)
Button(f1,text='CASH \nPAYMENT',font='Helvetica 15',command=lambda:r_frame(f4)).grid(row=7,column=0)
Label(f4,text=' ').grid(row=0,column=0)
Label(f4,text=' ').grid(row=1,column=0)
Label(f4,text=' ').grid(row=3,column=0)
Label(f4,text=' ').grid(row=2,column=0)
Label(f4,text='HAVE A NICE JOURNEY :)',font='Helvetica 20 bold').grid(row=7,column=0)
Label(f4,text=' ').grid(row=5,column=0)
Label(f4,text=' ').grid(row=6,column=0)
Label(f4,text='PLEASE GIVE CASH TO EMPLOYEE',font='Helvetica 13 ').grid(row=4,column=0)
Label(f4,text=' ').grid(row=2,column=1)
Label(f4,text=' ').grid(row=3,column=1)
Label(f4,text=' ').grid(row=4,column=1)
Button(f4,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)
Label(f2,text=' ').grid(row=1,column=0)
Label(f2,text=' ').grid(row=2,column=0)
Label(f2,text='PLEASE SWIPE YOUR CARD :)',font='Helvetica 30 bold ').grid(row=3,column=1)

def check(i,var1,var2):
    if((var1==1)&(var2==0)):
        import xlrd
        wb=xlrd.open_workbook(r'F:\python\shubh-python\idpass2.xls')
        s=wb.sheet_by_index(0)
        aa=s.col_values(0)
        bb=s.col_values(1)
        cc=s.col_values(2)
        dd=s.col_values(3)
        ee=s.col_values(4)
        ff=s.col_values(5)
        if(ff[i]>=10):
            import xlwt
            ws=xlwt.Workbook()
            s1=ws.add_sheet('one')
            ws.save(r'F:\python\shubh-python\idpass2.xls')
            ff[i]=ff[i]-10
            print(ff[i])
            
            for j in range(0,len(ff)):
                s1.write(j,5,ff[j])
                s1.write(j,0,aa[j])
                s1.write(j,1,bb[j])
                s1.write(j,2,cc[j])
                s1.write(j,3,dd[j])
                s1.write(j,4,ee[j])
            ws.save(r'F:\python\shubh-python\idpass2.xls')
            r_frame(f3)
            Label(f3,text=' ').grid(row=0,column=0)
            Label(f3,text=' ').grid(row=1,column=0)
            Label(f3,text=' ').grid(row=3,column=0)
            Label(f3,text=' ').grid(row=2,column=0)
            Label(f3,text='HAVE A NICE JOURNEY :)',font='Helvetica 20 bold').grid(row=4,column=0)
            Label(f3,text=' ').grid(row=5,column=0)
            Label(f3,text=' ').grid(row=6,column=0)
            Label(f3,text='Remaining balance is :- \n\t\t'+str(ff[i]),font='Helvetica 13 ').grid(row=7,column=0)
            Label(f3,text=' ').grid(row=8,column=0)
            Label(f3,text=' ').grid(row=9,column=0)
            Label(f3,text=' ').grid(row=10,column=0)
             #for motors when the code is applied on pi
            '''
            rp.setup(5,rp.IN)
            rp.setup(11,rp.OUT)
            rp.setup(13,rp.OUT)
            rp.setup(7,rp.IN)
            rp.output(11,1)
            rp.output(13,0)
            t.sleep(3)
            rp.output(11,0)
            rp.output(13,0)
            while(1):
                A=rp.imput(5)
                B=rp.imput(7)
                if((A==1)&(B==0)):
                    rp.output(11,0)
                    rp.output(13,1)
                    t.sleep(3)
                    rp.output(11,0)
                    rp.output(13,0)
                    break'''
            Button(f3,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)

            
           
        else:
            r_frame(f8)
            Label(f8,text=' ').grid(row=0,column=0)
            Label(f8,text=' ').grid(row=1,column=0)
            Label(f8,text=' ').grid(row=3,column=0)
            Label(f8,text=' ').grid(row=2,column=0)
            Label(f8,text='HAVE A NICE JOURNEY',font='Helvetica 20 bold').grid(row=7,column=0)
            Label(f8,text=' ').grid(row=5,column=0)
            Label(f8,text=' ').grid(row=6,column=0)
            Label(f8,text='INSUFFICIENT BALANCE\n\n  Give cash to employee',font='Helvetica 13 ').grid(row=4,column=0)
            Label(f8,text=' ').grid(row=8,column=0)
            Label(f8,text=' ').grid(row=9,column=0)
            Label(f8,text=' ').grid(row=10,column=0)
            Button(f8,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)

    elif((var2==1)&(var1==0)):
        import xlrd
        wb=xlrd.open_workbook(r'F:\python\shubh-python\idpass2.xls')
        s=wb.sheet_by_index(0)
        aa=s.col_values(0)
        bb=s.col_values(1)
        cc=s.col_values(2)
        dd=s.col_values(3)
        ee=s.col_values(4)
        ff=s.col_values(5)
        if(ff[i]>=15):
            import xlwt
            ws=xlwt.Workbook()
            s1=ws.add_sheet('one')
            ws.save(r'F:\python\shubh-python\idpass2.xls')
            ff[i]=ff[i]-15
            print(ff[i])
            
            for j in range(0,len(ff)):
                s1.write(j,5,ff[j])
                s1.write(j,0,aa[j])
                s1.write(j,1,bb[j])
                s1.write(j,2,cc[j])
                s1.write(j,3,dd[j])
                s1.write(j,4,ee[j])
            ws.save(r'F:\python\shubh-python\idpass2.xls')
            r_frame(f5)
            Label(f5,text=' ').grid(row=0,column=0)
            Label(f5,text=' ').grid(row=1,column=0)
            Label(f5,text=' ').grid(row=3,column=0)
            Label(f5,text=' ').grid(row=2,column=0)
            Label(f5,text='HAVE A NICE JOURNEY :)',font='Helvetica 20 bold').grid(row=4,column=0)
            Label(f5,text=' ').grid(row=5,column=0)
            Label(f5,text=' ').grid(row=6,column=0)
            Label(f5,text='Remaining balance is :- \n\t\t'+str(ff[i]),font='Helvetica 13 ').grid(row=7,column=0)
            Label(f5,text=' ').grid(row=8,column=0)
            Label(f5,text=' ').grid(row=9,column=0)
            Label(f5,text=' ').grid(row=10,column=0)
            #for motors when the code is applied on pi
            '''rp.setup(5,rp.IN)
            rp.setup(11,rp.OUT)
            rp.setup(13,rp.OUT)
            rp.setup(7,rp.IN)
            rp.output(11,1)
            rp.output(13,0)
            t.sleep(3)
            rp.output(11,0)
            rp.output(13,0)
            while(1):
                A=rp.imput(5)
                B=rp.imput(7)
                if((A==1)&(B==0)):
                    rp.output(11,0)
                    rp.output(13,1)
                    t.sleep(3)
                    rp.output(11,0)
                    rp.output(13,0)
                    break'''
            Button(f5,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)
        else:
            r_frame(f9)
            Label(f9,text=' ').grid(row=0,column=0)
            Label(f9,text=' ').grid(row=1,column=0)
            Label(f9,text=' ').grid(row=3,column=0)
            Label(f9,text=' ').grid(row=2,column=0)
            Label(f9,text='HAVE A NICE JOURNEY',font='Helvetica 20 bold').grid(row=7,column=0)
            Label(f9,text=' ').grid(row=5,column=0)
            Label(f9,text=' ').grid(row=6,column=0)
            Label(f9,text='INSUFFICIENT BALANCE\n\n  Give cash to employee',font='Helvetica 13 ').grid(row=4,column=0)
            Label(f9,text=' ').grid(row=8,column=0)
            Label(f9,text=' ').grid(row=9,column=0)
            Label(f9,text=' ').grid(row=10,column=0)
            Button(f9,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)

    elif((var1==1)&(var2==1)):
        r_frame(f6)
        Label(f6,text=' ').grid(row=0,column=0)
        Label(f6,text=' ').grid(row=1,column=0)
        Label(f6,text=' ').grid(row=3,column=0)
        Label(f6,text=' ').grid(row=2,column=0)
        Label(f6,text='INVALID INPUT',font='Helvetica 20 bold').grid(row=4,column=0)
        Label(f6,text=' ').grid(row=7,column=0)
        Label(f6,text=' ').grid(row=8,column=0)
        Label(f6,text=' ').grid(row=9,column=0)
        Button(f6,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)

    elif((var1==0)&(var1==0)):
        r_frame(f7)
        Label(f7,text=' ').grid(row=0,column=0)
        Label(f7,text=' ').grid(row=1,column=0)
        Label(f7,text=' ').grid(row=3,column=0)
        Label(f7,text=' ').grid(row=2,column=0)
        Label(f7,text='INVALID INPUT',font='Helvetica 20 bold').grid(row=4,column=0)
        Label(f7,text=' ').grid(row=8,column=0)
        Label(f7,text=' ').grid(row=9,column=0)
        Label(f7,text=' ').grid(row=10,column=0)
        Button(f7,text='Refresh',font='Helvetica 15',command=lambda:r_frame(f1)).grid(row=11,column=1)



