// <copyright file="OfficeHelper.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>02/05/2017 </date>
namespace Converters
{
    /// <summary>
    /// Class representing the office helper
    /// </summary>
    public static class OfficeHelper
    {
        /// <summary>
        /// Get the office installed
        /// </summary>
        /// <returns>The office installed</returns>
        public static IOffice GetOffice()
        {
            if (LibreOfficeHelper.IsInstalled())
            {
                return new LibreOffice();
            }

            if (MSOfficeHelper.IsInstalled())
            {
                return new MsOffice();
            }

            return new DefaultOffice();
        }
    }
}
