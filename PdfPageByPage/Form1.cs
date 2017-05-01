using Converters;
using iTextSharp.text.pdf;
using Spire.Pdf;
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

namespace PdfPageByPage
{
    public partial class PdfSplitter : Form
    {
        private int nbpages;
        private int progressPages;
        public PdfSplitter()
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
                if (Path.GetExtension(file) == FileExtension.PDF)
                {
                    this.Print(file);
                }
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

        void Print(string filepath)
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
                string outfile = filepath.Replace((System.IO.Path.GetFileName(filepath)), (System.IO.Path.GetFileName(filepath).Replace(FileExtension.PDF, string.Empty) + "_" + i.ToString()) + ext);
                outfile = outfile.Substring(0, dirPath.Length).Insert(dirPath.Length, "") + "\\" + Path.GetFileName(filepath).Replace(FileExtension.PDF, string.Empty) + "_" + i.ToString() + ext;
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
            this.ChangeMode(Color.FromArgb(197, 56, 39), Color.White, PdfPageByPage.Properties.Resources.drop);            
        }
        private void InSleep()
        {
            this.ChangeMode(Color.White, Color.FromArgb(197, 56, 39), PdfPageByPage.Properties.Resources.drag);
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
