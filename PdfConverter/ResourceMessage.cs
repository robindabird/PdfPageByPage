// <copyright file="ResourceMessage.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfConverter
{
    /// <summary>
    /// A class representing all the names of the UI messages
    /// </summary>
    public static class ResourceMessage
    {
        /// <summary>
        /// Get the resource name of the status when the calculation of the number of files begins
        /// </summary>
        public const string STATUS_CALCUL_NB_FILES = "status_calcul_nb_files";

        /// <summary>
        /// Get the resource name of the status when the conversion begins
        /// </summary>
        public const string STATUS_CONVERT_BEGIN = "status_convert_begin";

        /// <summary>
        /// Get the resource name of the status when the conversion is processing
        /// </summary>
        public const string STATUS_CONVERT_LOADING = "status_convert_loading";

        /// <summary>
        /// Get the resource name of the status when the conversion is completed
        /// </summary>
        public const string STATUS_CONVERT_END = "status_convert_end";

        /// <summary>
        /// Get the resource name of the status when no office is installed
        /// </summary>
        public const string STATUS_ERROR_OFFICE = "status_error_office";

        /// <summary>
        /// Get the resource name of the status when the calculation of the number of files is completed
        /// </summary>
        public const string STATUS_END_CALCUL_NB_FILES = "status_end_calcul_nb_files";

        /// <summary>
        /// Get the resource name of the status when there is no conversion
        /// </summary>
        public const string STATUS_NONE = "status_none";

        /// <summary>
        /// Get the resource name of the status when there is no documents
        /// </summary>
        public const string DOCUMENT_NONE = "document_none";
    }
}
