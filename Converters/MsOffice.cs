using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converters
{
    public class MsOffice : IOffice
    {
        public Microsoft.Office.Interop.Word.Document wordDocument { get; set; }
        public void ConvertToPdf(string filePath)
        {
            string fileExt = "." + Path.GetExtension(filePath);
            string pdfExt = ".pdf";
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string dirPath = Path.GetDirectoryName(filePath);
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            wordDocument = appWord.Documents.Open(filePath);
            wordDocument.ExportAsFixedFormat(dirPath + @"\" + fileName + pdfExt, WdExportFormat.wdExportFormatPDF);
            wordDocument.Close();
        }
    }
}
