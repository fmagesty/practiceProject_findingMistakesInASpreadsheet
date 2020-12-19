# Reference sheet: https://docs.google.com/spreadsheets/d/1jDZEdvSIh4TmZxccyy0ZXrH-ELlrwq8_YYiZrEOB4jg/edit#gid=289119951
# There is a mistake with the total bean count on this 'Bean Count' spreadsheet on one of its 15000 rows.
# This program finds the mistake and corrects it.

import ezsheets

# Load the spreadsheet.
print('Loading spreadsheet...')
ss = ezsheets.Spreadsheet('1Soy5EB5yBLyfkZJEdnq4J-PzFPR7NwPPl8XZYpcT4q0')
sheet = ss[0]
i = 2 # Is equal 2 because the first row is a header. So it checks from row 2 onwards.

# Evaluate if the column 'C' has the right value using a loop to check all cells. Also check if the cell is not a blank value.
print('Checking for errors...')
for row in sheet.getRows(): # Checks all rows.
    if sheet.getRow(i)[0] != '': # Checks if cell in the first column of each row is not a blank.
        if int(sheet.getRow(i)[0]) * int(sheet.getRow(i)[1]) != int(sheet.getRow(i)[2]):
            # Checks if a mistake is made on the total bean count.
            print('Error encountered.')
            # Corrects the mistake.
            sheet.getRow(i)[2] = str(int(sheet.getRow(i)[0]) * int(sheet.getRow(i)[1]))
            print(sheet.getRow(i)[2])
            print('Error corrected.')
    i += 1
# Saves the spreadsheet as a copy.
print('Saving a copy without errors...')
ss.downloadAsExcel('Copy of ' + str(ss.title) + '.xlsx')
print('Done.')