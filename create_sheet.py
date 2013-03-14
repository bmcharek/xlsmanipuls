#!bin/python
"""
Klein scriptje om een xls sheet te manipuleren.
"""
#TODO: Add specification for sheet number
import xlrd
import re
import xlwt
a_file = "Meevart_newMaildataMaart13.xls"
#"Adressen_mailings_wijkenbuurt_jan2013_bestversion.xls"
book = (xlrd.open_workbook(a_file))

nomatch = {}
email_adresss = []
email_adresss_dict = {}
sheet = book.sheet_by_index(1)
regstring = "[\-\_A-Za-z0-9]+[\.A-Za-z0-9]*@[\-A-Za-z0-9]+[\.a-zA-Z]+"
#import pdb; pdb.set_trace()

for counter in range(sheet.nrows):
    rowValue = sheet.row_values(counter, start_colx=0, end_colx=1)[0]
    try:
        matchs = re.findall(regstring, str.strip(str(rowValue)))
        if matchs:
            for match in matchs:
                email_adress = match
                if email_adress not in email_adresss:
                    email_adresss.append(str(email_adress))
                    if len(matchs) > 1:
                        full_rowValue = "Een mailinglijst"
                    else:
                        full_rowValue = str.strip(str(rowValue))
                    email_adresss_dict[str(email_adress)] = [full_rowValue,
                                                             str(email_adress).
                                                             split("@")[0]]
        else:
            nomatch[str(counter)] = rowValue
            #print "No match: ", rowValue, ", ",  counter
    except UnicodeEncodeError:
        matchs = re.findall(regstring, str.strip(rowValue.encode('ascii', 'replace')))
        for match in matchs:
            email_adress = match
            if email_adress not in email_adresss:
                email_adresss.append(str(email_adress))
                if len(matchs) > 1:
                    full_rowValue = "Een mailinglijst"
                else:
                    full_rowValue = (str.
                                     strip(rowValue.encode('ascii',
                                                            'replace')))
                email_adresss_dict[str(email_adress)] = ([full_rowValue,
                                                         str(email_adress).
                                                         split("@")[0]])
wb = xlwt.Workbook()
ws = wb.add_sheet("Adressen wijkenbuurt jan 2013")
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

i = 1
for email_adress in sorted(email_adresss, key=str.lower):
    #print "<"+email_adress+">,"
    sorted_email = "<"+email_adress+">,"
    try:
        ws.write(i, 1, email_adresss_dict[str(email_adress)][1])
        ws.write(i, 3, email_adresss_dict[str(email_adress)][0])
        full_name = str.strip(re.sub("[<]*"+regstring+"[>,;]*",
                 '',
                 email_adresss_dict[str(email_adress)][0]))
        if full_name:
            ws.write(i, 0, full_name)
        else:
            ws.write(i, 0, email_adresss_dict[str(email_adress)][1])
    except KeyError:
        print 'ooops'
    se_link = 'HYPERLINK("mailto:'+sorted_email + '";"'+sorted_email+'")'
    ws.write(i, 2, xlwt.Formula(se_link))
    i += 1
print "number of email:  ", len(email_adresss)
wb.save("adr2.xls")
