// <copyright file="FileHelper.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace Converters
{
    using System.IO;

    /// <summary>
    /// Class representing a file helper
    /// </summary>
    public static class FileHelper
    {
        /// <summary>
        /// Gets the file's information
        /// </summary>
        /// <param name="filePath">The path to the file</param>
        /// <returns>The information needed represented as a <see cref="FileCustom"/> object</returns>
        public static FileCustom GetFileInfo(string filePath)
        {
            string fileExt = Path.GetExtension(filePath);
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            string dirPath = Path.GetDirectoryName(filePath);
            return new FileCustom { Extension = fileExt, Name = fileName, Directory = dirPath };
        }
    }
}
