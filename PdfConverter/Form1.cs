using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using Spire.Doc;
using iTextSharp.text.pdf;

namespace PdfConverter
{
    public partial class PdfConvert : Form
    {
        private int nbpages;
        private int progressPages;
        public PdfConvert()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.DragLeave += new EventHandler(Form1_DragLeave);
            this.InSleep();
        }

        void Form1_DragLeave(object sender, EventArgs e)
        {
            this.InSleep();
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                this.InAction();
                e.Effect = DragDropEffects.Copy;
            }
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.ProcessNbPages(files);
            this.progressBar1.Value = 1;
            foreach (string file in files)
            {
                this.Print(file);
            }
            this.InSleep();
        }

        private void ProcessNbPages(string[] files)
        {
            this.nbpages = 0;
            this.progressPages = 0;
            if (files == null)
            {
                return;
            }
            this.statut.Text = "Début du calcul du nombre fichiers";
            this.statut.Refresh();
            foreach (string filepath in files)
            {
                this.nbpages += 1;
            }

            this.progressBar1.Minimum = 1;
            this.progressBar1.Maximum = this.nbpages;
            this.progressBar1.Step = 1;
            this.label1.Text = "0 / " + this.nbpages;
            this.label1.Refresh();
            this.statut.Text = "Fin du calcul du nombre de fichiers";
            this.statut.Refresh();
        }

        void Print(string filepath)
        {
            this.statut.Text = "Début de la conversion";
            this.statut.Refresh();
            this.document.Text = Path.GetFileName(filepath);
            this.document.Refresh();
            string fileExt = "." + Path.GetExtension(filepath);
            string pdfExt = ".pdf";
            string fileName = Path.GetFileNameWithoutExtension(filepath);
            string dirPath = Path.GetDirectoryName(filepath);
            Document doc = new Document();
            //Pass path of Word Document in LoadFromFile method  
            doc.LoadFromFile(filepath);
            //Pass Document Name and FileFormat of Document as Parameter in SaveToFile Method  
            doc.SaveToFile(dirPath + @"\" + fileName + pdfExt, FileFormat.PDF);
            //Launch Document  
            iTextSharp.text.pdf.PdfReader reader1 = new iTextSharp.text.pdf.PdfReader(dirPath + @"\" + fileName + pdfExt);
            
            string outfile = filepath.Replace((System.IO.Path.GetFileName(filepath)), (fileName.Replace(fileExt, "") + pdfExt));
            outfile = outfile.Substring(0, dirPath.Length).Insert(dirPath.Length, string.Empty) + "\\" + fileName + "_tmp" + pdfExt;
            reader1.RemoveUnusedObjects();
            PdfReader reader = null;
            reader = new iTextSharp.text.pdf.PdfReader(dirPath + @"\" + fileName.Replace(pdfExt, string.Empty) + pdfExt);
            reader.RemoveUnusedObjects();
            iTextSharp.text.Document doc2 = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(1));
            iTextSharp.text.pdf.PdfCopy pdfCpy = new iTextSharp.text.pdf.PdfCopy(doc2, new System.IO.FileStream(outfile, System.IO.FileMode.OpenOrCreate));
            doc2.Open();
            int pageCount = reader.NumberOfPages;
            for (int j = 2; j <= pageCount; j++)
            {
                iTextSharp.text.pdf.PdfImportedPage page = pdfCpy.GetImportedPage(reader1, j);
                pdfCpy.AddPage(page);
            }
            this.progressBar1.PerformStep();
            this.progressPages = this.progressPages + 1;
            this.label1.Text = this.progressPages + " / " + this.nbpages;
            this.label1.Refresh();
            this.statut.Text = "Fin de la conversion";
            this.statut.Refresh();
            doc.Close();
            doc2.Close();
            pdfCpy.Close();
            reader1.Close();
            reader.Close();
            File.Copy(outfile, dirPath + @"\" + fileName + pdfExt, true);
            File.Delete(outfile);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                // Read the files
                this.ProcessNbPages(openFileDialog1.FileNames);
                this.progressBar1.Value = 1;
                foreach (String file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    this.Print(file);
                }
            }

            this.InSleep();
        }
        private void InAction()
        {
            this.ChangeMode(Color.FromArgb(197, 56, 39), Color.White, PdfConverter.Properties.Resources.drop);            
        }
        private void InSleep()
        {
            this.ChangeMode(Color.White, Color.FromArgb(197, 56, 39), PdfConverter.Properties.Resources.drag);
        }

        private void ChangeMode(Color backColor, Color foreColor, Bitmap ressource)
        {
            this.pictureBox1.BackColor = backColor;
            this.pictureBox1.Image = ressource;
            this.label1.BackColor = backColor;
            this.label1.ForeColor = foreColor;
            this.document.BackColor = backColor;
            this.document.ForeColor = foreColor;
            this.statut.BackColor = backColor;
            this.statut.ForeColor = foreColor;
        }
    }
}
