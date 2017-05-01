// <copyright file="DefaultOffice.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    /// <summary>
    /// Class representing the office when no office is installed
    /// </summary>
    public class DefaultOffice : IOffice
    {
        /// <summary>
        /// Converts the file to a PDF file
        /// </summary>
        /// <param name="filePath">The full path of the file to be converted</param>
        public void ConvertToPdf(string filePath)
        {
            throw new NotInstalledOfficeException();
        }
    }
}
