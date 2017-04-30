# PdfPageByPage
PdfPageByPage splits a PDF document by page.

Example :
  You split a PDF file, named my_file.pdf containing 7 pages, the result will be 7 separated files named my_file_{page_number}.pdf in the same directory.
Third party libraries used :
  TextSharp : used to read the main file, and to create the split files.

## How to use ?
Launch the application, click to add files or drag and drop your pdf files.

# PdfConverter
PdfConvert converts docx, doc, txt, etc.. files into pdf files.
## This project uses the evaluation licence of spire.doc and is not meant to be used in a professionnal environnment.
Third party libraries used :
  Spire.Doc : used to convert documents
  TextSharp : used to read the converted file, and to create the final file. And to work around the message of evaluation added to the beginning of each converted files.

## How to use ?
Launch the application, click to add files or drag and drop your pdf files.
