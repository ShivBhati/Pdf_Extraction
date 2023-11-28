import os
import openpyxl
import pandas as pd
import fitz
from PyPDF2 import PdfReader

def extracting_pdf_pages(x,y,z):

    exdir = os.listdir(x)
    
    for e in exdir:
        exlfilename = os.path.join(x,e)
        with open(exlfilename, 'rb') as exl:
            df = pd.DataFrame()
            df = pd.read_excel(exl)
        exl.close
        openpdf(y,df,z)

def openpdf(y,df,z):
    exdir = os.listdir(y)
    for p in exdir:
        input_pdf = os.path.join(y,p)
        pdf_document = PdfReader(input_pdf)
        pdf_num_pages = len(pdf_document.pages)
        for index, row in df.iterrows():
            startpage = ''
            endpage = ''
            found = 'No'
            for page_num in range(pdf_num_pages):
                page = pdf_document.pages[page_num]
                text = page.extract_text()
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
    extracting_pdf_pages(open_file, pdf_path, extract_file_path)

if __name__ == "__main__":
    main()







