# Extract keywords from multiple PDF files, create a data frame, then export it to an .xlsx file.

## Import libraries
import pandas as pd
import os
import tabula 

## Function to extract information from a pdf document.
def ppt_toExcel():
    ## Take input from user
    print('Enter the path to the folder containing the PDF files:', end = '')
    path = input()

    print('Enter the path you want the excel file to be saved: ', end='')
    excel_path = input()

    folder = os.listdir(path) ## list all the files in the folder
    
    for item in folder: ## Loop through all the files in the folder
        base = os.path.basename(item) ## Get the file name
        if base.endswith('.pdf'): ## Check if the file is a pdf file
            pdfFileObj = open(item, 'rb') ## Open the file
            pdfReader1 = tabula.read_pdf(pdfFileObj, pages = 1) ## Read the file
            pdfReader2 = tabula.read_pdf(pdfFileObj, pages = 2) ## Read the file
            
            ## Combine for loops and existing file together
            df1 = pd.concat(pdfReader1, axis = 1)
            df2 = pd.concat(pdfReader2, axis = 1)
            df_com = pd.merge(df1, df2)

        pdfFileObj.close() ## Close the file
    
    df_new = pd.DataFrame(df_com) ## Create a data frame

    print("Document's keywords have been extracted successfully!") ## Print a message to the user

    df_new.to_excel(excel_path + '\Merged_file.xlsx', index = False) ## Export to an excel file

ppt_toExcel() ## Call the function
