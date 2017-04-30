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

namespace PdfConverter
{
    public partial class PdfConvert : Form
    {
        private int nbpages;
        private int progressPages;
        private bool officeNotInstalled = false;
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
            if (!this.officeNotInstalled)
            {
                this.statut.Text = "Début de la conversion";
                this.statut.Refresh();
                this.document.Text = Path.GetFileName(filepath);
                this.document.Refresh();
                Type officeType = Type.GetTypeFromProgID("Word.Application");
                bool isLibreOfficeInstalled = this.IsLibreOfficeInstalled();
                IOffice office;
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
                    this.statut.Text = "Conversion en cours";
                    this.statut.Refresh();
                    office.ConvertToPdf(filepath);
                }
                catch (NotInstalledOfficeException)
                {
                    this.officeNotInstalled = true;
                    this.statut.Text = "Veuillez installer Word ou LibreOffice";
                    MessageBox.Show("Veuillez installer Word ou LibreOffice");
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
