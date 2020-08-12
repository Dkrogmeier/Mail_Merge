import gspread #pip
import os
from oauth2client.service_account import ServiceAccountCredentials
import pprint #pip
from fpdf import FPDF  #http://fpdf.org/en/doc/cell.htm
from datetime import date
import time
start_time = time.time()

#Cell(float w [, float h [, string txt [, mixed border [, int ln [, string align [, boolean fill [, mixed link]]]]]]])


#-------------------------------------MESSAGE BOX--------------------------------------

import tkinter as tk
from tkinter import messagebox

root = tk.Tk()
root.withdraw()

messagebox.showinfo("Mail Merge", "Click OK to start, additional messages will appear")
messagebox.showinfo("Mail Merge", "Please close the Mail Merge PDF if its open, before clicking OK. ")



#---------------------------------GET INFO FROM PRODUCER SHEET------------------------------

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('Mail Merge Project.json', scope)
client = gspread.authorize(creds)
sheet = client.open('2020 07 FC Producer Quality Report SANDBOX').sheet1

prod_value = sheet.col_values(1)
prod_number = sheet.col_values(1)
del_date = sheet.col_values(2)
first_name = sheet.col_values(4)
last_name = sheet.col_values(6)
address = sheet.col_values(7)
CSZip = sheet.col_values(8)

weight = sheet.col_values(9)
weight += [''] * (len(last_name)-len(weight))
tank = sheet.col_values(10)
tank += [''] * (len(last_name)-len(tank))
BFV = sheet.col_values(11)
BFV += [''] * (len(last_name)-len(BFV))
PROV = sheet.col_values(12)
PROV += [''] * (len(last_name)-len(PROV))
OSV = sheet.col_values(13)
OSV += [''] * (len(last_name)-len(OSV))
SCCV = sheet.col_values(14)
SCCV += [''] * (len(last_name)-len(SCCV))
PIV = sheet.col_values(15)
PIV += [''] * (len(last_name)-len(PIV))
FRZV = sheet.col_values(16)
FRZV += [''] * (len(last_name)-len(FRZV))
MUNV = sheet.col_values(17)
MUNV += [''] * (len(last_name)-len(MUNV))
INHV = sheet.col_values(18)
INHV += [''] * (len(last_name)-len(INHV))
TEMPV = sheet.col_values(19)
TEMPV += [''] * (len(last_name)-len(TEMPV))



#---------------------------CREATE PDF------------------------------------

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'BI', 10)
        self.cell(20)
        self.cell(30,0, '2206 540th St SW, Kalona, IA 52247', 0, 1, 'C')  #0,\n
        self.ln(30)
        
    def footer(self):
        
        self.set_y(-10)
        self.set_font('Arial', 'B', 8)
        self.cell(190, 0, '', 1, 1, 'C')
        self.cell(0, 10, 'Generated:{}'.format(date.today()), 0, 0, 'C')

 
pdf = PDF()
#pdf.alias_nb_pages()

text = "Open Gates Group"
text2 = "Quality Systems"
text3 = "Quality Statement"
x = date.today()

pdf.set_font("Arial", 'B', size = 15)



#   ----------------------------------START INFO-------------------------------------
i = 1
l = 0
sizeOf = len(last_name)
while i < sizeOf:   #-----------LOOP INFO FOR SAME PERSON--------------
    if last_name[l].lower() == last_name[i].lower():
        pdf.cell(22, 5, ln = 0)
        pdf.cell(10, 5, txt = tank[i], ln = 0, align ='C')
        pdf.cell(30, 5, txt = del_date[i], ln = 0, align ='C')
        pdf.cell(17, 5, txt = weight[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = BFV[i], ln = 0, align ='C')
        pdf.cell(10, 5, txt = PROV[i], ln = 0, align ='C')
        pdf.set_text_color(255, 0, 0)
        pdf.cell(10, 5, txt = OSV[i], ln = 0, align ='C')
        pdf.set_text_color(0, 0, 0)
        pdf.cell(10, 5, txt = SCCV[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = PIV[i], ln = 0, align ='C')
        pdf.cell(9, 5, txt = FRZV[i], ln = 0, align ='C')
        pdf.cell(14, 5, txt = MUNV[i], ln = 0, align ='C')
        pdf.cell(10, 5, txt = INHV[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = TEMPV[i], ln = 1, align ='C')
        i += 1
        l += 1

        #-----------------------FINAL LINE PER PERSON--------------------------
        
        if last_name[l+1].lower() != last_name[l].lower():
            pdf.set_font("Times", 'B', size = 11)
            pdf.cell(22, 5, ln = 0)
            pdf.cell(156, 0, ln = 1, border = 1)
            
            pdf.cell(22, 5, ln = 0)
            pdf.cell(10, 5, txt = tank[l], ln = 0, align ='C')
            pdf.cell(30, 5, txt = del_date[l], ln = 0, align ='C')
            pdf.cell(17, 5, txt = weight[l], ln = 0, align ='C')
            pdf.cell(12, 5, txt = BFV[l], ln = 0, align ='C')
            pdf.cell(10, 5, txt = PROV[l], ln = 0, align ='C')
            pdf.set_text_color(255, 0, 0)
            pdf.cell(10, 5, txt = OSV[l], ln = 0, align ='C')
            pdf.set_text_color(0, 0, 0)
            pdf.cell(10, 5, txt = SCCV[l], ln = 0, align ='C')
            pdf.cell(12, 5, txt = PIV[l], ln = 0, align ='C')
            pdf.cell(9, 5, txt = FRZV[l], ln = 0, align ='C')
            pdf.cell(14, 5, txt = MUNV[l], ln = 0, align ='C')
            pdf.cell(10, 5, txt = INHV[l], ln = 0, align ='C')
            pdf.cell(12, 5, txt = TEMPV[l], ln = 1, align ='C')
            pdf.set_font("Arial", 'BU', size = 15)
            

#           ---------------------NEW PERSON INFO START---------------------

    else:
        #pdf.set_font("Arial", 'BU', size = 15)
        pdf.add_page()
        #pdf.image("Open Gates Logo Color.jpg", x = 60,y = 2, w = 90)
        #pdf.cell(190,8, ln = 1, align = "C", border = 1)
        pdf.set_font("Times", 'B', size = 14)
        pdf.cell(120, 4, ln = 0, align = 'L')
        pdf.set_font("Times", 'IB', size = 22)
        pdf.cell(70, 8, txt = text, ln = 1, align = 'L')
        pdf.set_font("Times", 'B', size = 14)
        pdf.cell(120, 5, txt = " {}".format(prod_number[i].zfill(5)), ln = 0, align = 'L')
        pdf.cell(70, 5, txt = text2, ln = 1, align = 'L')
        pdf.cell(120, 5, txt = " " + first_name[i].title() + " " + last_name[i].title(), ln = 0, align = 'L')
        pdf.cell(70, 5, txt = text3, ln = 1, align = 'L')
        pdf.cell(120, 5, txt = " " + address[i], ln = 0, align = 'L')
        pdf.cell(70, 5, txt = str(x), ln = 1, align = 'L')
        pdf.cell(120, 5, txt = " " + CSZip[i], ln = 0, align = 'L')
        pdf.set_font("Times", 'B', size = 12)

        pdf.set_font("Times", 'B', size = 12)

        pdf.cell(0, 20, ln = 1)
        pdf.cell(200, 5, txt = "Producer:  {}    {} {}".format(prod_number[i].zfill(5), first_name[i].title(), last_name[i].title()), ln = 1, align = 'L')
        pdf.cell(200, 5, txt = "Field Rep:    Phil Forbes", ln = 1, align = 'L')
        pdf.cell(0, 15, ln = 1)
        pdf.set_font("Times", 'BU', size = 13)
        pdf.cell(200, 6, txt = "Tank " + " Delivery Date " + "  Pounds  " + " BF " + " PRO " + " OS  " + " SCC  " + " PI " + " FRZ " + " MUN  " + " Inh " + " Temp", ln = 1, align = 'C')
        pdf.cell(22, 5, ln = 0)
        pdf.set_font("Times", 'B', size = 11)
        pdf.cell(10, 5, txt = tank[i], ln = 0, align ='C')
        pdf.cell(30, 5, txt = del_date[i], ln = 0, align ='C')
        pdf.cell(17, 5, txt = weight[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = BFV[i], ln = 0, align ='C')
        pdf.cell(10, 5, txt = PROV[i], ln = 0, align ='C')
        pdf.set_text_color(255, 0, 0)
        pdf.cell(10, 5, txt = OSV[i], ln = 0, align ='C')
        pdf.set_text_color(0, 0, 0)
        pdf.cell(10, 5, txt = SCCV[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = PIV[i], ln = 0, align ='C')
        pdf.cell(9, 5, txt = FRZV[i], ln = 0, align ='C')
        pdf.cell(14, 5, txt = MUNV[i], ln = 0, align ='C')
        pdf.cell(10, 5, txt = INHV[i], ln = 0, align ='C')
        pdf.cell(12, 5, txt = TEMPV[i], ln = 1, align ='C')
        i += 1
        l += 1
        
        
# save the pdf with name .pdf
nameMe = "Mailer " + str(x) + ".pdf"

try:
    pdf.output(nameMe)
    messagebox.showinfo("Mail Merge", "Finished! It took {} seconds".format(round(time.time()- start_time),2))
    
except (PermissionError) as e:
    #print(e)
    messagebox.showinfo("Mail Merge", "Please close previous PDF file and try again.")

os.startfile("Mailer " + str(x) + ".pdf")
