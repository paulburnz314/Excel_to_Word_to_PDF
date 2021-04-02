# Excel_to_Word_to_PDF
Automate importing Excel data into a Word template and then create a PDF.  Make one document or make 100!  

Windows 10 Requirements include:

  Install the latest version of python, I am now using 3.9.

  Install PyCharm Community 2020.3.4
  
  Pycharm will have to pip install:
  - comtypes
  - openpyxl
  - python-docx
  - docx-mailmerge
  
  Processed data will first become a word document and then that word document is closed.
  
  Then using comtypes.client starts a background version of word running:
  word = comtypes.client.CreateObject('Word.Application')
  
  Above is the slowest part of the the code.
    
  The code will combine labels with data to form a dictionary which is then appended to a list of dictionaries.
  
  The list of dicts can be used later to resort the data for other purposes.
  
