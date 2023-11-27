import os
import openpyxl
import pandas as pd
import fitz

def extracting_pdf_pages(x,y,z):

    exdir = os.listdir(x)
    df = pd.DataFrame()
    for e in exdir:
        with open(x, 'rb') as exl:
            df = openpyxl.load_workbook(exl,read_only=False)
        exl.close
        openpdf(y,df,z)

def openpdf(y,df,z):
    exdir = os.listdir(y)
    for p in exdir:
        input_pdf = os.path.join(y,p)
        pdf_document = fitz.open(input_pdf)
        for index, row in df.iterrows():
            startpage = ''
            endpage = ''
            found = 'No'
            for page_num in range(len(pdf_document)):
                page = pdf_document[page_num]
                text = page.get_text()
                if row['Invoice_Number'] in text and 'Statement' not in text:
                    if startpage == "": 
                        startpage =  page
                    else:
                        endpage = page
                    found = 'yes'
                else:
                    continue
            output_document = fitz.open()
            output_document.insert_pdf(pdf_document, from_page=startpage, to_page=endpage)
            output_path = os.path.join(z,row['Invoice_Number'])    
            output_document.save(output_path)
            output_document.close()
            if found == 'yes':
                break
            else:
                continue
        pdf_document.close()
        


        
    


def main():

    open_file = r'B:\Python\Git\Pdf_Extraction\excel'
    pdf_path = r'B:\Python\Git\Pdf_Extraction\pdf'
    extract_file_path = r'B:\Python\Git\Pdf_Extraction\extracted_pages'
    df = extracting_pdf_pages(open_file, pdf_path, extract_file_path)








