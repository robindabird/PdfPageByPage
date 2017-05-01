// <copyright file="LibreOfficeHelper.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>02/05/2017 </date>
namespace Converters
{
    using System;

    /// <summary>
    /// Class representing the LIBREOFFICE helper
    /// </summary>
    public static class LibreOfficeHelper
    {
        /// <summary>
        /// Check if LIBREOFFICE is installed
        /// </summary>
        /// <returns>True if LIBREOFFICE is installed, false otherwise</returns>
        public static bool IsInstalled()
        {
            string unoPath = string.Empty;
            bool installed = false;

            // access 32bit registry entry for latest LibreOffice for Current User
            Microsoft.Win32.RegistryKey hkcuView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry32);
            Microsoft.Win32.RegistryKey hkcuUnoInstallPathKey = hkcuView32.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false);
            if (hkcuUnoInstallPathKey != null && hkcuUnoInstallPathKey.ValueCount > 0)
            {
                unoPath = (string)hkcuUnoInstallPathKey.GetValue(hkcuUnoInstallPathKey.GetValueNames()[hkcuUnoInstallPathKey.ValueCount - 1]);
            }
            else
            {
                // access 32bit registry entry for latest LibreOffice for Local Machine (All Users)
                Microsoft.Win32.RegistryKey hklmView32 = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry32);
                Microsoft.Win32.RegistryKey hklmUnoInstallPathKey = hklmView32.OpenSubKey(@"SOFTWARE\LibreOffice\UNO\InstallPath", false);
                if (hklmUnoInstallPathKey != null && hklmUnoInstallPathKey.ValueCount > 0)
                {
                    installed = true;
                    unoPath = (string)hklmUnoInstallPathKey.GetValue(hklmUnoInstallPathKey.GetValueNames()[hklmUnoInstallPathKey.ValueCount - 1]);
                }
            }

            if (!Environment.GetEnvironmentVariable("PATH").Contains(unoPath))
            {
                Environment.SetEnvironmentVariable(
                    "PATH",
                    Environment.GetEnvironmentVariable("PATH") + @";" + unoPath,
                    EnvironmentVariableTarget.Process);
            }

            return installed;
        }
    }
}
