// <copyright file="MSOfficeHelper.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>02/05/2017 </date>
namespace Converters
{
    using System;

    /// <summary>
    /// Class representing the MSOFFICE helper
    /// </summary>
    public static class MSOfficeHelper
    {
        /// <summary>
        /// Checks if the word application is installed
        /// </summary>
        /// <returns>True if word is installed, false otherwise</returns>
        public static bool IsInstalled()
        {
            Type officeType = Type.GetTypeFromProgID(MSOfficeProperties.WORD_APPLICATION);
            return officeType != null;
        }
    }
}
