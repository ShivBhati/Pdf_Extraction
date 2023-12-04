import os
import openpyxl
import pandas as pd
from PyPDF2 import PdfReader
from PyPDF2 import PdfWriter



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
            pdf_writer = PdfWriter()
            for page_num in range( pdf_num_pages + 1):
                # print(f'the pdf {p} contains {page_num} pages')
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
    currentpath = os.path.dirname(os.path.abspath(__file__))
    open_file = os.path.join(currentpath,'excel')
    pdf_path = os.path.join(currentpath,'pdf')
    extract_file_path = os.path.join(currentpath,'extracted_pages')
    extracting_pdf_pages(open_file, pdf_path, extract_file_path)

if __name__ == "__main__":
    main()







