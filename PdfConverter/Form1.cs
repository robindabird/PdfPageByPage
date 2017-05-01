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
using Microsoft.Win32;
using Converters;
using System.Globalization;
using System.Globalization;
using System.Resources;
namespace PdfConverter
{
    public partial class PdfConvert : Form
    {
        private int nbpages;
        private int progressPages;
        private bool officeNotInstalled = false;
        ResourceManager ressourceManager;
        CultureInfo cultureInfo; 
        public PdfConvert()
        {
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            this.DragLeave += new EventHandler(Form1_DragLeave);
            this.InSleep();
            this.ressourceManager = new ResourceManager("PdfConverter.Ressource.Res", typeof(PdfConvert).Assembly);
            this.InitCultureInfo();
        }

        public void InitCultureInfo()
        {
            this.UpdateTextLabel(this.document, ressourceManager.GetString(RessourceMessage.DOCUMENT_NONE));
            this.UpdateTextLabel(this.statut, ressourceManager.GetString(RessourceMessage.STATUS_NONE));
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
                if (Path.GetExtension(file).ToLowerInvariant() != FileExtension.PDF)
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

            this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_CALCUL_NB_FILES));
            foreach (string filepath in files)
            {
                this.nbpages += 1;
            }

            this.progressBar1.Minimum = 1;
            this.progressBar1.Maximum = this.nbpages;
            this.progressBar1.Step = 1;
            this.UpdateTextLabel(this.label1, "0 / " + this.nbpages);
            this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_END_CALCUL_NB_FILES));
        }
        public void UpdateTextLabel(Label label, string text)
        {
            label.Text = text;
            label.Refresh();
        }
        
        void Print(string filepath)
        {
            if (!this.officeNotInstalled)
            {
                this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_CONVERT_BEGIN));
                this.UpdateTextLabel(this.document, Path.GetFileName(filepath));
                Type officeType = Type.GetTypeFromProgID("Word.Application");
                bool isLibreOfficeInstalled = this.IsLibreOfficeInstalled();
                IOffice office;
                isLibreOfficeInstalled = false;
                if (isLibreOfficeInstalled)
                {
                    office = new LibreOffice();
                }
                else if (officeType != null)
                {
                    office = new MsOffice();
                }
                else
                {
                    office = new DefaultOffice();
                }
                try
                {
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_CONVERT_LOADING));
                    office.ConvertToPdf(filepath);
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_CONVERT_END));
                    this.UpdateTextLabel(this.label1, ++this.progressPages + "/" + this.nbpages);
                    this.progressBar1.PerformStep();
                }
                catch (NotInstalledOfficeException)
                {
                    this.officeNotInstalled = true;
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(RessourceMessage.STATUS_ERROR_OFFICE));
                    MessageBox.Show(this.ressourceManager.GetString(RessourceMessage.STATUS_ERROR_OFFICE));
                }
            }
        }

        private bool IsLibreOfficeInstalled()
        {
            String unoPath = "";
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
                Environment.SetEnvironmentVariable("PATH",
                Environment.GetEnvironmentVariable("PATH") + @";" + unoPath, EnvironmentVariableTarget.Process);
            }

            return installed;
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
                    if (Path.GetExtension(file) != FileExtension.PDF)
                    {
                        this.Print(file);
                    }
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
