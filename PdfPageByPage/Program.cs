// <copyright file="Program.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfPageByPage
{
    using System;
    using System.Windows.Forms;

    /// <summary>
    /// The main entry of <see cref="PdfSplitter" /> application
    /// </summary>
    public static class Program
    {
        /// <summary>
        /// The main entry of the application
        /// </summary>
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new PdfSplitter());
        }
    }
}
