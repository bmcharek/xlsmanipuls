#!bin/python
"""
Klein scriptje om twee reeds gemanipuleerde xls te mergen
en te sorteren.
Kan natuurlijk uitgebreid worden tot een lijst van files.
"""
#TODO: Add specification for sheet number
import xlrd
import re
import xlwt
regstring = "[\-\_A-Za-z0-9]+[\.A-Za-z0-9]*@[\-A-Za-z0-9]+[\.a-z]+"
a_file = "Meevart_newMaildataMaart13.xls" #"adr.xls"
a_book = (xlrd.open_workbook(a_file))
a_sheet = a_book.sheet_by_index(0)
b_file = "adr2.xls"
b_book = (xlrd.open_workbook(b_file))
b_sheet = b_book.sheet_by_index(0)
email_adresss = []
email_fullnames = []
email_adresss_dict={}
rowValues = []
for counter in range(a_sheet.nrows):
    rowValues.append(a_sheet.row_values(counter, start_colx=0, end_colx=4))
for counter in range(b_sheet.nrows):
    rowValues.append(b_sheet.row_values(counter, start_colx=0, end_colx=4))

i = 0
for rowValue in rowValues:
    email_adress = rowValue[2]#mailadress
    if str(email_adress) not in email_adresss:
        ## if i is 11:
        ##     import pdb;pdb.set_trace()
        i+=1
        email_adresss.append(str(email_adress))
        email_fullname = rowValue[0]#fullname
        email_fullnames.append(str(email_fullname)+str(email_adress))
        rowValue.append(str(email_adress))
        email_adresss_dict[email_fullname+str(email_adress)] = rowValue 
        #    else:
        #print rowValue

        ## if "sichtman" in email_adress:
        ##     import pdb; pdb.set_trace()


wb = xlwt.Workbook()
ws = wb.add_sheet("Adressen wijkenbuurt maart 2013")
ws.col(0).width = 256 * max([len(row) for row in email_adresss])
ws.col(1).width = 256 * max([len(row) for row in email_adresss])
ws.col(2).width = 256 * max([len(row) for row in email_adresss_dict])
ws.col(3).width = 256 * max([len(row) for row in email_adresss_dict])


borders = xlwt.Borders()
borders.left = xlwt.Borders.DASHED
borders.right = xlwt.Borders.DASHED
borders.top = xlwt.Borders.DASHED
borders.bottom = xlwt.Borders.DASHED
borders.left_colour = 0x40
borders.right_colour = 0x40
borders.top_colour = 0x40
borders.bottom_colour = 0x40
style = xlwt.XFStyle()
style.borders = borders
styles = dict(
    bold='font: bold 1',
    italic='font: italic 1',
    # Wrap text in the cell
    wrap_bold='font: bold 1; align: wrap 1;',
    # White text on a blue background
    reversed='pattern: pattern solid, fore_color blue; font: color white;',
    # Light orange checkered background
    light_orange_bg='pattern: pattern fine_dots, fore_color white, back_color orange;',
    # Heavy borders
    bordered='border: top thick, right thick, bottom thick, left thick;',
    # 16 pt red text
    big_red='font: height 320, color red;',
)

ws.write(0, 0, 'Volle naam', style)
ws.write(0, 1, 'Gebruikers Naam', style)
ws.write(0, 2, 'Mail adres', style)
ws.write(0, 3, 'Oorspronkelijke cel inhoud', style)
i = 0
for email_fullname in sorted(email_fullnames, key=str.lower):
    i += 1
    ws.write(i, 0, str(email_adresss_dict[email_fullname][0]))#fullname
    print email_adresss_dict[email_fullname]
    ws.write(i, 1, str(email_adresss_dict[email_fullname][1]))#username
    ws.write(i, 2, str(email_adresss_dict[email_fullname][4]))#email
    ws.write(i, 3, str(email_adresss_dict[email_fullname][3]))#full length entry
    

print "number of email3:  ", len(email_adresss)
wb.save("adr3.xls")    
