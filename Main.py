
# launch on 10 July where i have seprate the whole 2a aside and whole PR Aside
# no oracle remarks
# final sheet will arrived after fp gstn difference
#




from tkinter.filedialog import askdirectory
import PyPDF2
import os

from tkinter import *
import  difflib, distance
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
import pandas as pd
from tkinter import messagebox
import threading
import time
import random
import numpy as np
from datetime import date
import re
#
# from datetime import datetime

import warnings
warnings.filterwarnings("ignore")


# a1=datetime.now()


datet = '2023-1-17'

today1=time.strftime("%Y-%m-%d")


while today1>datet:
    r1=messagebox.askretrycancel("INEFFICIENT MEMORY","THERE IS SOME ERROR OCCURRED")
    if r1==1:
        exit()
    else:
        exit()

else:
    pass

# now1=datetime.now()
global k
k=random.randint(1000,99999999)

global screen
screen=tk.Tk()
screen.config(bg='#1B998B')
screen.geometry('750x500')
screen.title('R-3.2 ')
l2=tk.Label(screen,text="CREATED BY \n RITESH VASWAL ",font=('bold',10),fg='#F8F1FF',bg='#1B998B')
l2.place(x=1000,y=5)

l1=tk.Label(screen,text=' DELHIVERY LIMITED ',font=('bold',25,UNDERLINE),fg='#F8F1FF',bg='#1B998B')
l1.place(x=20,y=20)

style = ttk.Style()
style.map("C.TButton",
        foreground=[('pressed', 'pink'), ('active', 'blue')],
        background=[('pressed', '!disabled', 'black'), ('active', 'white')])



line1="                                                                                                                                                                    "

l2 = tk.Label(screen, text=line1, font=('bold', 15, UNDERLINE), fg='#F8F1FF', bg='#1B998B')
l2.place(x=20, y=280)



Button_checkinv = ttk.Button(screen, text='PR BOOK ', style="C.TButton", command=lambda: open2inv())
Button_checkinv.place(x=20, y=320)

global entry3

entry3=Text(screen,height=1,width=60,bd=5,borderwidth=5)
entry3.place(x=150,y=320)


Button_open=ttk.Button(screen,text='OPEN',style="C.TButton",command=lambda :open1())
Button_open.place(x=20,y=100)

global entry1
entry1=Text(screen,height=1,width=60,bd=5,borderwidth=5)
entry1.place(x=120,y=100)


def open1():

    global path1
    path1=askopenfilename()


    entry1.insert(END,path1)

    convert1=ttk.Button(screen,text='CONVERT',style="C.TButton",command=lambda :convertfile())
    convert1.place(x=20,y=150)

    extract1 = ttk.Button(screen, text='EXTRACT', style="C.TButton",command=lambda:extract2())
    extract1.place(x=20, y=220)

    line1="                                                                                                                                                                    "

    l2 = tk.Label(screen, text=line1, font=('bold', 15, UNDERLINE), fg='#F8F1FF', bg='#1B998B')
    l2.place(x=20, y=280)

    Button_checkinv = ttk.Button(screen, text='check book', style="C.TButton", command=lambda: open2inv())
    Button_checkinv.place(x=20, y=320)

    global progress

    progress = Progressbar(screen, orient = HORIZONTAL,
              length = 300, mode = 'determinate')
    progress.place(x=20,y=180)


def convertfile():

    try:
        global df
        df=pd.read_excel(path1)
        progress['value'] = 10
        screen.update_idletasks()


        df=df.fillna(0)

        df = df.loc[(df['REC'] == 'N') & (df['Vendor Name'] != 'DELHIVERY PRIVATE LIMITED') & (
                    df['Vendor Name'] != 'DELHIVERY LIMITED') & (df['Vendor Name'] != 'DELHIVERY  LIMITED') & (
                                df['Reverse Charge'] != 'Y')]


        convert_dict = {'FP GSTN': str, 'GSTN Supplier': str, 'Vendor Name': str, 'Invoice Number': str,
                        'Invoice Date': str, 'Rate': str,
                        'FP GST': str, 'GSTN Supplier.1': str, 'Invoice Number.1': str, 'Invoice Date.1': str,
                        'Rate.1': str}  # not in books-2a and not not in 2a-books

        df = df.astype(convert_dict)
        pivot2a = df.groupby(['FP GSTN', 'GSTN Supplier', 'Vendor Name', 'Invoice Number', 'Invoice Date', 'Rate'],
                             as_index=False)[['Taxable Value', 'IGST', 'CGST', 'SGST']].mean()
        pivot2 = df.groupby(['FP GST', 'GSTN Supplier.1', 'Invoice Number.1', 'Invoice Date.1', 'Rate.1'], as_index=False)[
            ['Taxable Value.1', 'IGST.1', 'CGST.1', 'SGST.1']].mean()

        global df3,df4
        df3 = pd.DataFrame(pivot2a)
        df4 = pd.DataFrame(pivot2)
        df3 = df3[(df3['Taxable Value'] != 0)]
        df4 = df4[df4['Taxable Value.1'] != 0]



    #
#
#
        t1=threading.Thread(target=mapping1)
        t2=threading.Thread(target=mapping2)
        t3=threading.Thread(target=mapping3)

        t1.start()
        t2.start()
        t3.start()

    except Exception as e:
        messagebox.askretrycancel("FILTER",e)

#
# # =======================================================


def mapping1():
    try:

        df31 = df3.copy()
        df41 = df4.copy()

        num1 = []

        for i in df31['Invoice Number']:

            try:
                c1 = re.findall(r'\d+', str(i))
                res = int("".join(map(str, c1)))
                res = str(res)
                num1.append(res)

            except:
                num1.append(i)
                pass

        df31['num1'] = num1

        pan1 = []
        for index1, row1 in df31.iterrows():
            n2 = str(row1['GSTN Supplier'][:-3])  # change
            pan1.append(n2)

        df31['PAN_NO'] = pan1

        df31['MERGE1'] = df31['PAN_NO'].map(str) + df31['num1']

        # =========================================

        num2 = []

        for i in df41['Invoice Number.1']:

            try:
                c1 = re.findall(r'\d+', str(i))
                res = int("".join(map(str, c1)))
                res = str(res)
                num2.append(res)

            except:
                num2.append(i)
                pass

        df41['num2'] = num2

        pan2 = []
        for index2, row2 in df41.iterrows():
            n21 = str(row2['GSTN Supplier.1'][:-3])  # change
            pan2.append(n21)

        df41['PAN_NO1'] = pan2

        df41['MERGE1'] = df41['PAN_NO1'].map(str) + df41['num2']

        global df5_map_1

        df5_map_1 = pd.merge(df31,
                            df41,
                            on='MERGE1',
                            how='inner')


        df5_map_1 = df5_map_1.drop(['num1',
                                  'PAN_NO', 'MERGE1', 'num2', 'PAN_NO1'], axis=1)
        df5_map_1['STATUS'] = 'MAP-1'
        df5_map_1.to_csv(f'MAP-1-{k}.csv')

        tk.Label(screen, text="* MAP-1 HAS BEEN CREATED", font=('bold', 20), fg='#F8F1FF', bg='#1B998B').place(
            x=20, y=350)
        progress['value'] = 40

    except Exception as e:
        messagebox.askretrycancel('ERROR-MAP-1',e)




# =============================================
def mapping2():

    try:
        df5=df3.copy()
        df6=df4.copy()

        tk.Label(screen,text="* MAP-2 SHEET IS RUNNING",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=350)
        # =================================
        l2=[]
        for i in df5['Invoice Number']:

            if "-2021-2022" in i:
                k0=i.replace("-2021-2022","")
                l2.append(k0)

            elif "/2021-2022" in i:
                k01=i.replace("/2021-2022","")
                l2.append(k01)

            elif "\2021-2022" in i:
                k02=i.replace("\2021-2022","")
                l2.append(k02)


            elif "-21-22" in i:
                k1=i.replace("21-22","")
                l2.append(k1)

            elif "/21-22" in i:
                k12=i.replace("/21-22","")
                l2.append(k12)

            elif "\21-22" in i:
                k13=i.replace("\21-22","")
                l2.append(k13)



            elif "-2021-22" in i:
                k3=i.replace("-2021-22","")
                l2.append(k3)

            elif "/2021-22" in i:
                k31=i.replace("/2021-22","")
                l2.append(k31)

            elif "\2021-22" in i:
                k32=i.replace("\2021-22","")
                l2.append(k32)



            elif "-20-21" in i:
                k4=i.replace("-20-21","")
                l2.append(k4)

            elif "\20-21" in i:
                k41=i.replace("\20-21","")
                l2.append(k41)

            elif "/20-21" in i:
                k42=i.replace("/20-21","")
                l2.append(k42)


            elif "-2020-2021" in i:
                k5=i.replace("-2020-2021","")
                l2.append(k5)

            elif "\2020-2021" in i:
                k51=i.replace("\2020-2021","")
                l2.append(k51)

            elif "/2020-2021" in i:
                k52=i.replace("/2020-2021","")
                l2.append(k52)

            elif "-21-2022" in i:
                k6=i.replace("-21-2022","")
                l2.append(k6)

            elif "/21-2022" in i:
                k61=i.replace("/21-2022","")
                l2.append(k61)

            elif "\21-2022" in i:
                k62=i.replace("\21-2022","")
                l2.append(k62)

            elif "-2022/23" in i:
                k7=i.replace("-2022/23","")
                l2.append(k7)

            elif "/2022/23" in i:
                k71=i.replace("/2022/23","")
                l2.append(k71)

            elif "\2022/23" in i:
                k72=i.replace("\2022/23","")
                l2.append(k72)



            elif "/2022-23" in i:
                k8=i.replace("/2022-23","")
                l2.append(k8)

            elif "-2022-23" in i:
                k81=i.replace("-2022-23","")
                l2.append(k81)

            elif "\2022-23" in i:
                k83=i.replace("\2022-23","")
                l2.append(k83)



            elif "-2022-2023" in i:
                k9=i.replace("-2022-2023","")
                l2.append(k9)

            elif "/2022-2023" in i:
                k91=i.replace("/2022-2023","")
                l2.append(k91)

            elif "\2022-2023" in i:
                k92=i.replace("\2022-2023","")
                l2.append(k92)



            elif "-22-23" in i:
                k10=i.replace("-22-23","")
                l2.append(k10)

            elif "\22-23" in i:
                k101=i.replace("\22-23","")
                l2.append(k101)

            elif "/22-23" in i:
                k102=i.replace("/22-23","")
                l2.append(k102)


            elif "-2021/22" in i:
                k11=i.replace("-2021/22","")
                l2.append(k11)

            elif "\2021/22" in i:
                k111=i.replace("\2021/22","")
                l2.append(k111)

            elif "/2021/22" in i:
                k112=i.replace("/2021/22","")
                l2.append(k112)




            elif "-22/23" in i:
                k12=i.replace("-21/22","")
                l2.append(k12)

            elif "/22/23" in i:
                k121=i.replace("/21/22","")
                l2.append(k121)

            elif "\22/23" in i:
                k122=i.replace("\21/22","")
                l2.append(k122)




            elif "-21-22" in i:
                k13=i.replace("-21-22","")
                l2.append(k13)

            elif "/21-22" in i:
                k131=i.replace("/21-22","")
                l2.append(k131)

            elif "\21-22" in i:
                k132=i.replace("\21-22","")
                l2.append(k132)


            #     ===========




            else:
                l2.append(i)

        df5['invoiceno1']=l2




        # ===========================

        merge1=[]
        num1=[]

        for i in df5['invoiceno1']:

            try:
                c1=re.findall(r'\d+',str(i))
                res=int("".join(map(str, c1)))
                res=str(res)
                num1.append(res)

            except:
                num1.append(i)
                pass



        df5['num1']=num1



        pan=[]
        for index3,row3 in df5.iterrows():
            n11=str(row3['GSTN Supplier'][:-3])              #change
            pan.append(n11)

        df5['PAN_NO']=pan


        for index1,row1 in df5.iterrows():
            m1=str(row1['PAN_NO'])+str(row1['num1'])      #change
            merge1.append(m1)

        df5['MERGE1']=merge1
        # df5.to_excel(f'df5-{k}.xlsx')

        try:

            # =========================================
            l22=[]
            for i in df6['Invoice Number.1']:


                if "-2021-2022" in i:
                    k0=i.replace("-2021-2022","")
                    l22.append(k0)

                elif "/2021-2022" in i:
                    k01=i.replace("/2021-2022","")
                    l22.append(k01)

                elif "\2021-2022" in i:
                    k02=i.replace("\2021-2022","")
                    l22.append(k02)


                elif "-21-22" in i:
                    k1=i.replace("21-22","")
                    l22.append(k1)

                elif "/21-22" in i:
                    k12=i.replace("/21-22","")
                    l22.append(k12)

                elif "\21-22" in i:
                    k13=i.replace("\21-22","")
                    l22.append(k13)



                elif "-2021-22" in i:
                    k3=i.replace("-2021-22","")
                    l22.append(k3)

                elif "/2021-22" in i:
                    k31=i.replace("/2021-22","")
                    l22.append(k31)

                elif "\2021-22" in i:
                    k32=i.replace("\2021-22","")
                    l22.append(k32)



                elif "-20-21" in i:
                    k4=i.replace("-20-21","")
                    l22.append(k4)

                elif "\20-21" in i:
                    k41=i.replace("\20-21","")
                    l22.append(k41)

                elif "/20-21" in i:
                    k42=i.replace("/20-21","")
                    l22.append(k42)


                elif "-2020-2021" in i:
                    k5=i.replace("-2020-2021","")
                    l22.append(k5)

                elif "\2020-2021" in i:
                    k51=i.replace("\2020-2021","")
                    l22.append(k51)

                elif "/2020-2021" in i:
                    k52=i.replace("/2020-2021","")
                    l22.append(k52)

                elif "-21-2022" in i:
                    k6=i.replace("-21-2022","")
                    l22.append(k6)

                elif "/21-2022" in i:
                    k61=i.replace("/21-2022","")
                    l22.append(k61)

                elif "\21-2022" in i:
                    k62=i.replace("\21-2022","")
                    l22.append(k62)

                elif "-2022/23" in i:
                    k7=i.replace("-2022/23","")
                    l22.append(k7)

                elif "/2022/23" in i:
                    k71=i.replace("/2022/23","")
                    l22.append(k71)

                elif "\2022/23" in i:
                    k72=i.replace("\2022/23","")
                    l22.append(k72)



                elif "/2022-23" in i:
                    k8=i.replace("/2022-23","")
                    l22.append(k8)

                elif "-2022-23" in i:
                    k81=i.replace("-2022-23","")
                    l22.append(k81)

                elif "\2022-23" in i:
                    k83=i.replace("\2022-23","")
                    l22.append(k83)



                elif "-2022-2023" in i:
                    k9=i.replace("-2022-2023","")
                    l22.append(k9)

                elif "/2022-2023" in i:
                    k91=i.replace("/2022-2023","")
                    l22.append(k91)

                elif "\2022-2023" in i:
                    k92=i.replace("\2022-2023","")
                    l22.append(k92)



                elif "-22-23" in i:
                    k10=i.replace("-22-23","")
                    l22.append(k10)

                elif "\22-23" in i:
                    k101=i.replace("\22-23","")
                    l22.append(k101)

                elif "/22-23" in i:
                    k102=i.replace("/22-23","")
                    l22.append(k102)


                elif "-2021/22" in i:
                    k11=i.replace("-2021/22","")
                    l22.append(k11)

                elif "\2021/22" in i:
                    k111=i.replace("\2021/22","")
                    l22.append(k111)

                elif "/2021/22" in i:
                    k112=i.replace("/2021/22","")
                    l22.append(k112)




                elif "-22/23" in i:
                    k12=i.replace("-21/22","")
                    l22.append(k12)

                elif "/22/23" in i:
                    k121=i.replace("/21/22","")
                    l22.append(k121)

                elif "\22/23" in i:
                    k122=i.replace("\21/22","")
                    l22.append(k122)




                elif "-21-22" in i:
                    k13=i.replace("-21-22","")
                    l22.append(k13)

                elif "/21-22" in i:
                    k131=i.replace("/21-22","")
                    l22.append(k131)

                elif "\21-22" in i:
                    k132=i.replace("\21-22","")
                    l22.append(k132)


                #     ===========




                else:
                    l22.append(i)






            df6['invoiceno2']=l22



            # =======================================
            num2=[]
            merge2=[]
            for i1 in df6['invoiceno2']:
                try:

                    c2=re.findall(r'\d+',str(i1))
                    res1=int("".join(map(str, c2)))
                    res1=str(res1)
                    num2.append(res1)
                except:
                    num2.append(i1)
                    pass




        except:
            pass

        df6['num2']=num2
        # print(len(num2))

        # ****************************************************************


        pan2=[]
        for index4,row4 in df6.iterrows():
            n2=str(row4['GSTN Supplier.1'][:-3])               #change
            pan2.append(n2)

        df6['PAN_NO']=pan2


        for index2,row2 in df6.iterrows():
            m2=str(row2['PAN_NO'])+str(row2['num2'])                  #change
            merge2.append(m2)

        df6['MERGE1']=merge2

        global df6_map_2



        df6_map_2 = pd.merge(df5,
                              df6,
                              on ='MERGE1',
                              how ='inner')

        df6_map_2.drop(['num1','num2','MERGE1',"invoiceno1","invoiceno2"], axis=1, inplace=True)
        df6_map_2['STATUS']='MAP-2'
        df6_map_2.to_csv(f'map-2-{k}.csv')


        tk.Label(screen,text="* MAP-2 SHEET HAS BEEN CREATED",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=400)
        messagebox.showinfo('MAP-2','MAP-2 SHEET HAS BEEN CREATED')




    except Exception as e:
        messagebox.askokcancel(" MAP-2",e)
        tk.Label(screen,text=f"ERROR=={e}",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=400)

def mapping3():

    try:
        df5=df3.copy()
        df6=df4.copy()
        # tk.Label(screen,text="* MAP-3 SHEET IS RUNNING",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=450)

        pan=[]
        for index3,row3 in df5.iterrows():
            n11=str(row3['GSTN Supplier'][:-3])              #change
            pan.append(n11)

        df5['PAN_NO']=pan


        merge3=[]
        for index3,row3 in df5.iterrows():
            n1=str(row3['PAN_NO'])+str(row3['Taxable Value'])               #change
            merge3.append(n1)

        df5['MERGE2']=merge3


        pan2=[]
        for index4,row4 in df6.iterrows():
            n2=str(row4['GSTN Supplier.1'][:-3])               #change
            pan2.append(n2)

        df6['PAN_NO']=pan2


        merge4=[]
        for index4,row4 in df6.iterrows():
            n2=str(row4['PAN_NO'])+str(row4['Taxable Value.1'])               #change
            merge4.append(n2)
        df6['MERGE2']=merge4

        global df7_map_2

        df7_map_2= pd.merge(df5,
                              df6,
                              on ='MERGE2',
                              how ='inner')

        df7_map_2.drop(['MERGE2'], axis=1, inplace=True)
        df7_map_2['STATUS']='SUGGESTED'

        df7_map_2.to_csv(f'map-3-{k}.csv')



        # =========================================

        # ============================================================

        tk.Label(screen,text="* MAP-3 SHEET HAS BEEN CREATED",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=450)
        messagebox.showinfo("MAP-3-",'MAP-3 SHEET HAS BEEN CREATED')
        progress['value'] = 95
        screen.update_idletasks()



    except Exception as e:
        messagebox.askokcancel('MAP-3',e)
        tk.Label(screen,text=F"ERROR=={e}",font=('bold',20),fg='#F8F1FF',bg='#1B998B').place(x=20,y=500)



def extract2():
    r1=messagebox.askyesnocancel('EXTRACTION','please check ,Have all 3 map file completed')
    if r1==1:
        try:
            df_final=pd.concat([df5_map_1,df6_map_2,df7_map_2],ignore_index=True)
            # df_final.to_csv(f'{k}-check1.csv')

            df_final = df_final.astype(str)
            df_final['IGST.1']=pd.to_numeric(df_final['IGST.1'])
            df_final['CGST.1'] = pd.to_numeric(df_final['CGST.1'])
            df_final['SGST.1'] = pd.to_numeric(df_final['SGST.1'])


            remarks = ['fp gstn diff' if row1['FP GSTN'] != row1['FP GST'] else 'clear' for index1, row1 in df_final.iterrows()]
            df_final['REMARKS'] = remarks

            df_final = df_final[df_final['REMARKS'] == 'clear']

            bkcheck = []
            for index3, row3 in df_final.iterrows():
                if row3['FP GST'][:2] != row3['GSTN Supplier.1'][:2] and row3['IGST.1']== 0:
                    bkcheck.append('booking error')
                elif row3['FP GST'][:2] == row3['GSTN Supplier.1'][:2] and row3['CGST.1'] == 0 and row3['SGST.1']== 0:
                    bkcheck.append('booking error')
                else:

                    bkcheck.append('correct booking')

            df_final['book-check'] = bkcheck

            per = []

            for index5,row5  in df_final.iterrows():
                diffl = difflib.SequenceMatcher(None, row5['Invoice Number'], row5['Invoice Number.1']).ratio()
                # lev = Levenshtein.ratio(row5['Invoice Number'], row5['Invoice Number.1'])
                sor = 1 - distance.sorensen(row5['Invoice Number'], row5['Invoice Number.1'])
                jac = 1 - distance.jaccard(row5['Invoice Number'], row5['Invoice Number.1'])

                per1 = (diffl+ sor + jac) / 3 * 100
                #     print(per)
                per.append(per1)

            df_final['Match Percent']=per

            # df_final.to_csv(f'{k}-check.csv')

            concat1 = []
            for index2, row2 in df_final.iterrows():
                try:
                    e1 = (str(row2['Rate']) + str(row2['Rate.1']) + row2['FP GSTN'] + row2['GSTN Supplier'] + row2[
                        'Invoice Number'] + row2['FP GST'] + row2['GSTN Supplier.1'] + row2['Invoice Number.1'])
                    concat1.append(e1)
                except:
                    pass

            df_final['concat'] = concat1

            df_final.drop_duplicates(subset=['concat'], keep='first', inplace=True)

            df_final=df_final.drop(['PAN_NO_x','PAN_NO_y','concat','REMARKS'],axis=1)
            # print('yes====')


            df_final.to_csv(f'final-{k}.csv')

            messagebox.showinfo('EXTRACTION SUCCESS','FINAL SHEET ARRIVED')


        except Exception as e:
            print(e)
            messagebox.showinfo('WAITING','PLEASE WAIT UNTILL ALL FILES ARRIVED')
            pass

    else:
        pass

# ==============================================================================

# check book


def open2inv():
    global path2
    # messagebox.showinfo(' Information ', "choose your file")
    path2 = askopenfilename()




    entry3.insert(END, path2)

    convert2 = ttk.Button(screen, text='SELECT DIRECTORY', style="C.TButton", command=lambda: directory1())
    convert2.place(x=20, y=370)

    # extract1 = ttk.Button(screen, text='EXTRACT', style="C.TButton", command=lambda: extract2())
    # extract1.place(x=20, y=220)


    # Button_checkinv = ttk.Button(screen, text='check book', style="C.TButton", command=lambda: open2inv())
    # Button_checkinv.place(x=20, y=320)



def directory1():
    global path3
    path3=askdirectory()

    entry4 = Text(screen, height=1, width=60, bd=5, borderwidth=5)
    entry4.place(x=150, y=370)

    entry4.insert(END, path3)

    convert3 = ttk.Button(screen, text='Start Checking', style="C.TButton", command=lambda: pdf1())
    convert3.place(x=20, y=430)

# ======================================================================================

# from here we searching invoices on the directory list and invoice number

def pdf1():
    try:
        dir_list = os.listdir(f"{path3}")
    # print(dir_list)

        df_p=pd.read_excel(path2)

        for i in range(0, len(df_p['Invoice Number.1'])):
            for j in dir_list:
                k1=df_p['Invoice Number.1'][i]
                if str(k1) in j:
                    df_p.loc[i, "path-1"] = f'{j}'


        df_p['path-1'] = df_p['path-1'].fillna(0)
        p1 = df_p[df_p['path-1'] != 0]
    # print("p1====\n",p1)

        for index1, row in p1.iterrows():

        # print("path=====",row['path'])
            pdfFileObj = open(rf"{path3}/{row['path-1']}", 'rb')
            pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        # creating a page object
            pageObj = pdfReader.getPage(0)

        # extracting text from page
            global text1
            text1 = pageObj.extractText()
        # print(text1)
            pdfFileObj.close()

        # =====
            y1=[]
            y = re.findall('\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z\d]{1}[Z]{1}[A-Z\d]{1}', text1)
            y1.append(y)



            if str(row['Invoice Number.1']) in text1:
                p1.loc[index1, 'Invoice Num-check'] = 'True'
            else:
                p1.loc[index1, 'Invoice Num-check'] = 'False'

            if row['FP GST'] in text1:
                p1.loc[index1, 'FP GSTN-CHECK'] = 'True'

            else:
                p1.loc[index1, 'FP GSTN-CHECK'] = 'False'

            if row['GSTN Supplier.1'] in text1:
                p1.loc[index1, 'GSTN Supplier-check'] = 'True'

            else:
                p1.loc[index1, 'GSTN Supplier-check'] = 'False'

            if f"{str(round(row['Rate.1']))}%" in text1:
                p1.loc[index1, 'Rate-check'] = 'True'

            elif f"{str(round(row['Rate.1']))} %" in text1:
                p1.loc[index1, 'Rate-check'] = 'True'
            else:
            # print(f"{str(round(row['Rate.1']))}%")
                p1.loc[index1, 'Rate-check'] = 'False'


        p1.to_csv(f'c-booking-{k}.csv')
        messagebox.showinfo('Success','File arrived')
    except EXCEPTION as e:
        messagebox.showinfo('ERROR',e)
# ===============================================
screen.mainloop()
#
# # b=datetime.now()
# # print(b-a1)
