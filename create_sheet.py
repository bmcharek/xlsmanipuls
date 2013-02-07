import xlwt

DATA = (("The Essential Calvin and Hobbes", 1988,),
        ("The Authoritative Calvin and Hobbes", 1990,),
        ("The Indispensable Calvin and Hobbes", 1992,),
        ("Attack of the Deranged Mutant Killer Monster Snow Goons", 1992,),
        ("The Days Are Just Packed", 1993,),
        ("Homicidal Psycho Jungle Cat", 1994,),
        ("There's Treasure Everywhere", 1996,),
        ("It's a Magical World", 1996,),)

wb = xlwt.Workbook()
ws = wb.add_sheet("My Sheet")

# Add headings with styling and froszen first row
heading_xf = xlwt.easyxf('font: bold on; align: wrap on, vert centre, horiz center')
headings = ['Bill Watterson Books', 'Year Published']
rowx = 0
ws.set_panes_frozen(True) # frozen headings instead of split panes
ws.set_horz_split_pos(rowx+1) # in general, freeze after last heading row
ws.set_remove_splits(True) # if user does unfreeze, don't leave a split there
for colx, value in enumerate(headings):
    ws.write(rowx, colx, value, heading_xf)

for i, row in enumerate(DATA):
    for j, col in enumerate(row):
        ws.write(i, j, col)
ws.col(0).width = 256 * max([len(row[0]) for row in DATA])
wb.save("myworkbook.xls")
