// <copyright file="FileCustom.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    /// <summary>
    /// Class representing all the file's information needed to process the conversion to PDF
    /// </summary>
    public class FileCustom
    {
        /// <summary>
        /// Gets or sets the the extension
        /// </summary>
        public string Extension { get; set; }

        /// <summary>
        /// Gets or sets the name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the directory path
        /// </summary>
        public string Directory { get; set; }
    }
}
