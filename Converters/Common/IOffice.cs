// <copyright file="IOffice.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    /// <summary>
    /// Interface representing an office
    /// </summary>
    public interface IOffice
    {
        /// <summary>
        /// Converts the file to a PDF file
        /// </summary>
        /// <param name="filePath">The full path of the file to be converted</param>
        void ConvertToPdf(string filePath);
    }
}
