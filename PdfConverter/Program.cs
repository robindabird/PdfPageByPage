// <copyright file="Program.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfConverter
{
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// Class representing the main entry of the <see cref="PdfConvert"/> application
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The main entry of the <see cref="PdfConvert"/> application
        /// </summary>
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new PdfConvert());
        }
    }
}
