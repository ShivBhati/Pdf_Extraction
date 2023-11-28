PDF Page Extraction Tool

Overview:
This Python script is designed to extract specific pages from PDF files based on information provided in an Excel file. 
It is useful for scenarios where you have a collection of PDFs and an Excel sheet containing details such as Invoice Numbers, 
and you want to extract pages from the PDFs that match the specified criteria.

Features:
Excel Data Extraction: The script reads data from Excel files using the pandas library to create a DataFrame containing relevant information.

PDF Page Extraction: 
It iterates through the PDF files, searching for pages that contain the specified Invoice Number and meet certain criteria.
Extracted pages are saved into new PDF files.

Usage:
Install Dependencies: Make sure you have the required dependencies installed. You can install them using the following:

Author
Shivam Bhati