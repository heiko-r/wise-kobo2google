# Python script to create a PDF copy of responses and upload it to Google Drive

import pdfkit
from io import BytesIO

# Enable for verbose debug logging (disabled by default)
g_EnableDebugMsg = False


'''
Function to create a PDF with the copy of responses.
Returns a file handler to the created in-memory PDF file.
'''
def create_pdf(html):
    # Create PDF file in memory
    if g_EnableDebugMsg:
        pdfoptions = {}
    else:
        pdfoptions = {'quiet': ''}
    pdfdata = pdfkit.from_string(html, False, options=pdfoptions)
    return BytesIO(pdfdata)
