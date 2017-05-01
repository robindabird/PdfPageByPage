// <copyright file="MSOffice.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    using Microsoft.Office.Interop.Word;

    /// <summary>
    /// Class representing the office of Microsoft Word
    /// </summary>
    public class MsOffice : IOffice
    {
        /// <summary>
        /// Converts the file to a PDF file
        /// </summary>
        /// <param name="filePath">The full path of the file to be converted</param>
        public void ConvertToPdf(string filePath)
        {
            string pdfExt = FileExtension.PDF;
            Microsoft.Office.Interop.Word.Document wordDocument;
            FileCustom fileCustom = FileHelper.GetFileInfo(filePath);
            Microsoft.Office.Interop.Word.Application appWord = new Microsoft.Office.Interop.Word.Application();
            wordDocument = appWord.Documents.Open(filePath);
            wordDocument.ExportAsFixedFormat(fileCustom.Directory + @"\" + fileCustom.Name + pdfExt, WdExportFormat.wdExportFormatPDF);
        }
    }
}
