import comtypes.client  # for saving a word file as a pdf
import time     # for delay to avoid erros
import openpyxl, pprint     # for getting data from excel
from mailmerge import MailMerge     # for dumping values into word document where keys match mergefields

wrd_to_pdf = 17 #code for saving as a pdf in word
wb = openpyxl.load_workbook('Surcharge_example.xlsx') # excel file that holds raw data exported from database
surcharge_amount = [0.14, 0.13, 0.70, 5.08, 0.14]   # was [0.14, 0.13, 0.70, 5.08, 0.14]
threshold_amount = [200, 200, 15, 4, 100]     # was [200,200,15,4,100]
pollutant_id = ['TSS','CBOD','NH3N','TPhos','O&G']  #pollutants
siu_list = ['PPI','FFI','SCE','NCC','CBC']      #facility IDs
sheet1 = wb['Export1']      # name of the sheet with the data
sheet2 = wb['Contact2']     # name of the sheet with facility and contact information
Contact1 = sheet2['A2':'A6']
column_start = 3
list_value = 0
surcharge_part = 0
merge_field_labels = ['Contact_Name', 'User_Name', 'User_Address', 'Contact_Title', 'User_Code', 'Contact2',
                      'Contact3', 'Month_Year', 'Flow', 'TSSppm', 'TSS_Charge', 'CBODppm', 'CBOD_Charge',
                      'NH3Nppm', 'NH3N_Charge', 'TPhosppm', 'TPhos_Charge', 'O&Gppm', 'O&G_Charge','Total_Surcharge']
merge_match_results = []
total_surcharge = 0
users = siu_list[0]


def replaceMultiple(mainString, toBeReplaces, newString):   # Iterate over the strings to be replaced
    for elem in toBeReplaces:   # Check if string is in the main string
        if elem in mainString:  # Replace the string
            mainString = mainString.replace(elem, newString)
    return mainString

wb2 = openpyxl.Workbook()   # might as well dump the data processed into a new spreadsheet
sheet3 = wb2.active
sheet3.title = '2020 Summary'   # change as needed
sheet3.append(merge_field_labels)


for row_num in range(2,5):   # which rows do you want to process, 12 months would be range(2,13)
    month_year = sheet1.cell(row=row_num, column=1).value   # month and year always in column 1
    month_year = replaceMultiple(month_year, [','],'') # get the month and year but remove the comma and ...

    for user in siu_list:   # each user has six columns of data including flow, tss, cbod, nh3, tp and o&g
        doc_template = 'Surcharge_mailmerge_example.docx'
        document_1 = MailMerge(doc_template)
        flow = sheet1.cell(row=row_num, column = column_start - 1).value
        for column_num in range(1, 8):
            # fill the list merge_match_results with excel 'Contact2' sheet data
            merge_match_results.append(sheet2.cell(row=(siu_list.index(user))+2, column=column_num).value)

        merge_match_results.append(month_year)
        flow = f'{flow:.6f}'    # format flow in millions of gallons and maintain 6 decimal places
        merge_match_results.append(str(flow))
        document_title = month_year + '_' + user + '_surcharges.docx'   # name of the word doc to be generated
        document_title2 = month_year + '_' + user + '_surcharges.pdf'   # name of the pdf to be generated
        for column_num in range(column_start,column_start + 5):     # at start range is 3 to 8, note stops after column 7
            pollutant = sheet1.cell(row=row_num, column=column_num).value
            merge_match_results.append(str(pollutant))
            checkstr = isinstance(pollutant, str)   # start checking for <, >, empty space or None values
            if checkstr is True:
                pollutant = replaceMultiple(pollutant, ['>', '<'], '')
                if pollutant == ' ':
                    pollutant = 0
                else:
                    pollutant = float(pollutant)    # data is cleaned up and converted to a number
            if pollutant == None:
                pollutant = 0
            if pollutant > threshold_amount[list_value]:    # no surcharge if value is below threshold, no credit for low values either
                flow = float(flow)
                surcharge_part = round(flow * 8.34*(surcharge_amount[list_value]*(pollutant-threshold_amount[list_value])),2)
            total_surcharge = total_surcharge + surcharge_part
            str_surcharge_part = f'{surcharge_part:.2f}'
            merge_match_results.append(str_surcharge_part)
            list_value = list_value + 1
            surcharge_part = 0
        str_total_surcharge = f'{total_surcharge:.2f}'
        merge_match_results.append(str_total_surcharge)
        sheet3.append(merge_match_results)      #dump data into summary excel sheet

        # combine the two lists into a dictionary so you can pass it to the word document
        merge_dict = {merge_field_labels[i]: merge_match_results[i] for i in range(len(merge_field_labels))}
        surcharge_code = month_year + '_' + user + ' = '
        result_file = open('surcharges_dictionary.py', 'a')    # store the combined keys and values as a a seperate python file
        result_file.write("\n")
        result_file.write(surcharge_code +pprint.pformat(merge_dict))
        result_file.close()
        document_1.merge_templates([merge_dict], "page_break")
        # document_1.merge_pages([merge_dict])   NOTE if I was going to combine all the pages into 1 document I would use this
        document_1.write(document_title)
        document_1.close()
        #taking the just written document as in_file so it can be reopened and saved as a pdf
        in_file = 'C:\\Users\\paul\\Documents\\python docs\\python projects paul\\autosurcharge2\\' + document_title
        out_file = 'C:\\Users\\paul\\Documents\\python docs\\python projects paul\\autosurcharge2\\' + document_title2

        # creating COM object
        word = comtypes.client.CreateObject('Word.Application')
        # word.Visible = True       kind of annoying having every document window pop up
        time.sleep(1) # this slows down the file generation but avoids errors
        doc = word.Documents.Open(in_file)    # in_file is the word document just created
        doc.SaveAs(out_file, FileFormat=wrd_to_pdf)     # out_file makes the pdf
        doc.Close()
        #word.Visible = False

        #set the column_start to grab the next set of columns on the row
        column_start = column_start + 6
        total_surcharge = 0     # reset values just in case previous data is still in there
        list_value = 0
        merge_match_results.clear()
        merge_dict.clear()
        doc_template = ''

        wb2.save(filename='Surcharge_Summary.xlsx')
        # after going through one user's pollutants, move on to the next, once all are complete, drop to next row/month
    column_start = 3    #it is time to move down to the next row so set which column to start with

word.Quit()


