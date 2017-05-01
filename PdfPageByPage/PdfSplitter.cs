// <copyright file="PdfSplitter.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfPageByPage
{
    using System;
    using System.Drawing;
    using System.IO;
    using System.Windows.Forms;
    using Converters;
    using iTextSharp.text.pdf;

    /// <summary>
    /// Class representing the form for the PDFSPLITTER application
    /// </summary>
    public partial class PdfSplitter : Form
    {
        /// <summary>
        /// Gets or sets the number of pages to split
        /// </summary>
        private int nbpages;

        /// <summary>
        /// Get or sets the progress of pages split
        /// </summary>
        private int progressPages;

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfSplitter"/> class
        /// </summary>
        public PdfSplitter()
        {
            this.InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(this.PdfSplitter_DragEnter);
            this.DragDrop += new DragEventHandler(this.PdfSplitter_DragDrop);
            this.DragLeave += new EventHandler(this.PdfSplitter_DragLeave);
            this.InSleep();
        }

        /// <summary>
        /// Processed after the users leaves the drag action
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PdfSplitter_DragLeave(object sender, EventArgs e)
        {
            this.InSleep();
        }

        /// <summary>
        /// Processed when the user drags files onto the application
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PdfSplitter_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                this.InAction();
                e.Effect = DragDropEffects.Copy;
            }
        }

        /// <summary>
        /// Processed when the user drops some files
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PdfSplitter_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.ProcessNbPages(files);
            this.progressBar1.Value = 1;
            foreach (string file in files)
            {
                if (Path.GetExtension(file) == FileExtension.PDF)
                {
                    this.Split(file);
                }
            }

            this.InSleep();
        }

        /// <summary>
        /// Counts the number of pages to split
        /// </summary>
        /// <param name="files">The list of files to split</param>
        private void ProcessNbPages(string[] files)
        {
            this.nbpages = 0;
            this.progressPages = 0;
            if (files == null)
            {
                return;
            }

            this.statut.Text = "Début du calcul du nombre total de pages";
            this.statut.Refresh();
            foreach (string filepath in files)
            {
                PdfReader reader = null;
                int pageCount = 0;
                string dirPath = Path.GetDirectoryName(filepath);
                System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
                reader = new iTextSharp.text.pdf.PdfReader(filepath);
                reader.RemoveUnusedObjects();
                pageCount = reader.NumberOfPages;
                this.nbpages += pageCount;
            }

            this.progressBar1.Minimum = 1;
            this.progressBar1.Maximum = this.nbpages;
            this.progressBar1.Step = 1;
            this.label1.Text = "0 / " + this.nbpages;
            this.label1.Refresh();
            this.statut.Text = "Fin du calcul du nombre total de pages";
            this.statut.Refresh();
        }

        /// <summary>
        /// Splits the PDF file
        /// </summary>
        /// <param name="filepath">The file to split</param>
        private void Split(string filepath)
        {
            this.statut.Text = "Début de la division";
            this.statut.Refresh();
            this.document.Text = Path.GetFileName(filepath);
            this.document.Refresh();
            PdfReader reader = null;
            int currentPage = 1;
            int pageCount = 0;
            string dirPath = Path.GetDirectoryName(filepath);
            System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
            reader = new iTextSharp.text.pdf.PdfReader(filepath);
            reader.RemoveUnusedObjects();
            pageCount = reader.NumberOfPages;
            string ext = System.IO.Path.GetExtension(filepath);
            for (int i = 1; i <= pageCount; i++)
            {
                this.statut.Text = "En cours de division " + i + " / " + pageCount;
                this.statut.Refresh();
                iTextSharp.text.pdf.PdfReader reader1 = new iTextSharp.text.pdf.PdfReader(filepath);
                string outfile = filepath.Replace(System.IO.Path.GetFileName(filepath), (System.IO.Path.GetFileName(filepath).Replace(FileExtension.PDF, string.Empty) + "_" + i.ToString()) + ext);
                outfile = outfile.Substring(0, dirPath.Length).Insert(dirPath.Length, string.Empty) + "\\" + Path.GetFileName(filepath).Replace(FileExtension.PDF, string.Empty) + "_" + i.ToString() + ext;
                reader1.RemoveUnusedObjects();
                iTextSharp.text.Document doc = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(currentPage));
                iTextSharp.text.pdf.PdfCopy pdfCpy = new iTextSharp.text.pdf.PdfCopy(doc, new System.IO.FileStream(outfile, System.IO.FileMode.OpenOrCreate));
                doc.Open();
                for (int j = 1; j <= 1; j++)
                {
                    iTextSharp.text.pdf.PdfImportedPage page = pdfCpy.GetImportedPage(reader1, currentPage);
                    pdfCpy.AddPage(page);
                    currentPage += 1;
                }

                this.progressBar1.PerformStep();
                this.progressPages = this.progressPages + 1;
                this.label1.Text = this.progressPages + " / " + this.nbpages;
                this.label1.Refresh();
                this.statut.Text = "Fin de la division " + i + " / " + pageCount;
                this.statut.Refresh();
                doc.Close();
                pdfCpy.Close();
                reader1.Close();
                reader.Close();
            }
        }

        /// <summary>
        /// Processed after the user clicks onto the picture
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PictureBox1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                this.ProcessNbPages(openFileDialog1.FileNames);
                this.progressBar1.Value = 1;
                foreach (string file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    this.Split(file);
                }
            }

            this.InSleep();
        }

        /// <summary>
        /// Updates all the UI colors to make them appear as processing
        /// </summary>
        private void InAction()
        {
            this.ChangeMode(Color.FromArgb(197, 56, 39), Color.White, PdfPageByPage.Properties.Resources.drop);            
        }

        /// <summary>
        /// Updates all the UI colors to make them appear as not processing
        /// </summary>
        private void InSleep()
        {
            this.ChangeMode(Color.White, Color.FromArgb(197, 56, 39), PdfPageByPage.Properties.Resources.drag);
        }

        /// <summary>
        /// Changes the back color, fore color and the bitmap picture
        /// </summary>
        /// <param name="backColor">The back color of the application</param>
        /// <param name="foreColor">The fore color of the label</param>
        /// <param name="resource">The bitmap resource</param>
        private void ChangeMode(Color backColor, Color foreColor, Bitmap resource)
        {
            this.pictureBox1.BackColor = backColor;
            this.pictureBox1.Image = resource;
            this.label1.BackColor = backColor;
            this.label1.ForeColor = foreColor;
            this.document.BackColor = backColor;
            this.document.ForeColor = foreColor;
            this.statut.BackColor = backColor;
            this.statut.ForeColor = foreColor;
        }
    }
}
