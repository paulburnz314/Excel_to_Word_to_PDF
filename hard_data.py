SURCHARGE_AMOUNT = (0.14, 0.13, 0.70, 5.08, 0.14)
THRESHOLD_AMOUNT = (200, 200, 15, 4, 100)
SIU_LIST = ('PPI', 'FFI', 'SCE', 'NCC', 'CBC')  # facility IDs
POLLUTANT_ID = ('TSS', 'CBOD', 'NH3N', 'TPhos')  # pollutants
SIU_FULL_NAME = ('Poultry Processing Inc', 'Franklin Foods Inc ', 'Super Chicken Express',
                 'Notorious Chicken Co', 'Cardboard Box Co')
MERGE_FIELD_LABELS = (
        'Contact_Name', 'User_Name', 'User_Address', 'Contact_Title', 'User_Code',
        'Contact2', 'Contact3', 'Month_Year', 'Flow',
        'TSS_ppm', 'TSS_over', 'TSS_load', 'TSS_amt',
        'CBOD_ppm', 'CBOD_over', 'CBOD_load', 'CBOD_amt',
        'NH3N_ppm', 'NH3N_over', 'NH3N_load', 'NH3N_amt',
        'TP_ppm', 'TP_over', 'TP_load', 'TP_amt',
        'O&G_ppm', 'O&G_over', 'O&G_load', 'O&G_amt', 'Total_Surcharge'
    )
WRD_TO_PDF = 17  # code for saving as a pdf in word
PROJECT_FOLDER_OUTPUT = "C:\\Users\\paul\\Documents\\python docs\\python projects paul\\excel-to-word-to-pdf\\output\\"
