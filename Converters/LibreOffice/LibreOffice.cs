// <copyright file="LibreOffice.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    using System.IO;
    using unoidl.com.sun.star.frame;
    using unoidl.com.sun.star.lang;
    using unoidl.com.sun.star.util;

    /// <summary>
    /// Class representing the office of LIBREOFFICE
    /// </summary>
    public class LibreOffice : IOffice
    {
        /// <summary>
        /// Converts the file to a PDF file
        /// </summary>
        /// <param name="filePath">The full path of the file to be converted</param>
        public void ConvertToPdf(string filePath)
        {
            string pdfExt = FileExtension.PDF;
            FileCustom fileCustom = FileHelper.GetFileInfo(filePath);
            var xLocalContext = uno.util.Bootstrap.bootstrap();
            var xRemoteFactory = (XMultiServiceFactory)xLocalContext.getServiceManager();
            XComponentLoader loader = (XComponentLoader)xRemoteFactory.createInstance(LibreOfficeProperties.FRAME_DESKTOP);
            var pv = new unoidl.com.sun.star.beans.PropertyValue[1];
            pv[0] = new unoidl.com.sun.star.beans.PropertyValue();
            pv[0].Name = LibreOfficeProperties.HIDDEN;
            pv[0].Value = new uno.Any(true);
            XComponent xComponent;
            xComponent = loader.loadComponentFromURL("file:///" + filePath.Replace('\\', '/'), LibreOfficeProperties.BLANK, 0, pv);
            var xStorable = (XStorable)xComponent;
            pv[0].Name = LibreOfficeProperties.FILTER_NAME;
            switch (fileCustom.Extension.ToLowerInvariant())
            {
                case FileExtension.XLS:
                case FileExtension.XLSX:
                case FileExtension.ODS:
                    pv[0].Value = new uno.Any(LibreOfficeProperties.CALC_PDF_EXPORT);
                    break;
                default:
                    pv[0].Value = new uno.Any(LibreOfficeProperties.WRITER_PDF_EXPORT);
                    break;
            }

            xStorable.storeToURL("file:///" + (fileCustom.Directory + @"\" + fileCustom.Name + pdfExt).Replace('\\', '/'), pv);
            var xClosable = (XCloseable)xComponent;
            xClosable.close(true);
        }
    }
}
