using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using unoidl.com.sun.star.frame;
using unoidl.com.sun.star.lang;
using unoidl.com.sun.star.util;

namespace Converters
{
    public class LibreOffice : IOffice
    {
        public void ConvertToPdf(string filePath)
        {
            string fileExt = Path.GetExtension(filePath);
            string pdfExt = FileExtension.PDF;
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string dirPath = Path.GetDirectoryName(filePath);
            var xLocalContext = uno.util.Bootstrap.bootstrap();
            var xRemoteFactory = (XMultiServiceFactory)xLocalContext.getServiceManager();
            XComponentLoader loader = (XComponentLoader)xRemoteFactory.createInstance(LibreOfficeProperties.FRAME_DESKTOP);
            var pv = new unoidl.com.sun.star.beans.PropertyValue[1];
            pv[0] = new unoidl.com.sun.star.beans.PropertyValue();
            pv[0].Name = LibreOfficeProperties.HIDDEN ;
            pv[0].Value = new uno.Any(true);
            XComponent xComponent;
            xComponent = loader.loadComponentFromURL("file:///" + filePath.Replace('\\', '/'), LibreOfficeProperties.BLANK, 0, pv);
            var xStorable = (XStorable)xComponent;
            pv[0].Name = LibreOfficeProperties.FILTER_NAME;
            switch (fileExt.ToLowerInvariant())
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
            xStorable.storeToURL("file:///" + (dirPath + @"\" + fileName + "." + pdfExt).Replace('\\', '/'), pv);
            var xClosable = (XCloseable)xComponent;
            xClosable.close(true);
        }
    }
}
