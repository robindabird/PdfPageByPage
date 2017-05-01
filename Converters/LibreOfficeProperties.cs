// <copyright file="LibreOfficeProperties.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    /// <summary>
    /// Class representing all the properties of LIBREOFFICE
    /// </summary>
    public static class LibreOfficeProperties
    {
        /// <summary>
        /// Get the PDF property to export documents
        /// </summary>
        public const string WRITER_PDF_EXPORT = "writer_pdf_Export";

        /// <summary>
        /// Get the PDF property to export spreadsheets
        /// </summary>
        public const string CALC_PDF_EXPORT = "calc_pdf_Export";

        /// <summary>
        /// Get the blank property
        /// </summary>
        public const string BLANK = "_blank";

        /// <summary>
        /// Get the desktop property
        /// </summary>
        public const string FRAME_DESKTOP = "com.sun.star.frame.Desktop";

        /// <summary>
        /// Get the hidden property
        /// </summary>
        public const string HIDDEN = "Hidden";

        /// <summary>
        /// Get the filter nam property
        /// </summary>
        public const string FILTER_NAME = "FilterName";
    }
}
