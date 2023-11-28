import os
import openpyxl
import pandas as pd
import fitz
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter

pdf_writer = PdfWriter()

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
    for index, row in df.iterrows():
        for p in exdir:
            input_pdf = os.path.join(y,p)
            pdf_document = PdfReader(input_pdf)
            pdf_num_pages = len(pdf_document.pages)
            found = 'No'
            for page_num in range(0, pdf_num_pages+1):
                print(f'the pdf {p} contains {page_num} pages')
                page = pdf_document.pages[page_num]
                text = page.extract_text()
                if row['Invoice_Number'] in text and 'PLEASE PAY FROM THIS STATEMENT' not in text:
                    pdf_writer.add_page(page)
                    found = 'yes'
                else:
                    continue
            if found == 'yes':
                output_path = os.path.join(z,row['Invoice_Number']+".pdf")    
                with open(output_path,'wb') as output:
                    pdf_writer.write(output)
                break
            else:
                continue
        

def main():  
    open_file = r'B:\Python\Git\Pdf_Extraction\excel'
    pdf_path = r'B:\Python\Git\Pdf_Extraction\pdf'
    extract_file_path = r'B:\Python\Git\Pdf_Extraction\extracted_pages'
    extracting_pdf_pages(open_file, pdf_path, extract_file_path)

if __name__ == "__main__":
    main()







