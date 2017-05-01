// <copyright file="PdfConvert.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfConverter
{
    using System;
    using System.Drawing;
    using System.IO;
    using System.Resources;
    using System.Windows.Forms;
    using Converters;

    /// <summary>
    /// Class representing the form of the PDFCONVERT application
    /// </summary>
    public partial class PdfConvert : Form
    {
        /// <summary>
        /// Gets or sets the number of files to convert
        /// </summary>
        private int numberFiles;

        /// <summary>
        /// Gets or sets the number of files already converted
        /// </summary>
        private int progressFiles;

        /// <summary>
        /// Gets or sets a value
        /// </summary>
        private bool officeNotInstalled = false;

        /// <summary>
        /// Gets or sets the resource manager to process UI messages
        /// </summary>
        private ResourceManager ressourceManager;

        /// <summary>
        /// Initializes a new instance of the <see cref="PdfConvert"/> class
        /// </summary>
        public PdfConvert()
        {
            this.InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(this.PdfConvert_DragEnter);
            this.DragDrop += new DragEventHandler(this.PdfConvert_DragDrop);
            this.DragLeave += new EventHandler(this.PdfConvert_DragLeave);
            this.InSleep();
            this.ressourceManager = new ResourceManager(ResourceManagerProperties.RESOURCE, typeof(PdfConvert).Assembly);
            this.InitLabel();
        }

        /// <summary>
        /// Initializes the labels
        /// </summary>
        private void InitLabel()
        {
            this.UpdateTextLabel(this.document, this.ressourceManager.GetString(ResourceMessage.DOCUMENT_NONE));
            this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_NONE));
        }

        /// <summary>
        /// Processed after the users leaves the drag action
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PdfConvert_DragLeave(object sender, EventArgs e)
        {
            this.InSleep();
        }

        /// <summary>
        /// Processed when the user drags files onto the application
        /// </summary>
        /// <param name="sender">The object which sends the event</param>
        /// <param name="e">The arguments</param>
        private void PdfConvert_DragEnter(object sender, DragEventArgs e)
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
        private void PdfConvert_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            this.ProcessNbFiles(files);
            this.progressBar1.Value = 1;
            foreach (string file in files)
            {
                if (Path.GetExtension(file).ToLowerInvariant() != FileExtension.PDF)
                {
                    this.Converts(file);
                }
            }

            this.InSleep();
        }

        /// <summary>
        /// Counts the number of files to convert
        /// </summary>
        /// <param name="files">The list of files to convert</param>
        private void ProcessNbFiles(string[] files)
        {
            this.numberFiles = 0;
            this.progressFiles = 0;
            if (files == null)
            {
                return;
            }

            this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_CALCUL_NB_FILES));
            foreach (string filepath in files)
            {
                this.numberFiles += 1;
            }

            this.progressBar1.Minimum = 1;
            this.progressBar1.Maximum = this.numberFiles;
            this.progressBar1.Step = 1;
            this.UpdateTextLabel(this.label1, "0 / " + this.numberFiles);
            this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_END_CALCUL_NB_FILES));
        }

        /// <summary>
        /// Routine to help update the text in a label
        /// </summary>
        /// <param name="label">The label which the text will be updated</param>
        /// <param name="text">The text</param>
        private void UpdateTextLabel(Label label, string text)
        {
            label.Text = text;
            label.Refresh();
        }
        
        /// <summary>
        /// Converts the file to a PDF file
        /// </summary>
        /// <param name="filepath">The file to be converted</param>
        private void Converts(string filepath)
        {
            if (!this.officeNotInstalled)
            {
                this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_CONVERT_BEGIN));
                this.UpdateTextLabel(this.document, Path.GetFileName(filepath));
                IOffice office = OfficeHelper.GetOffice();
                try
                {
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_CONVERT_LOADING));
                    office.ConvertToPdf(filepath);
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_CONVERT_END));
                    this.UpdateTextLabel(this.label1, ++this.progressFiles + "/" + this.numberFiles);
                    this.progressBar1.PerformStep();
                }
                catch (NotInstalledOfficeException)
                {
                    this.officeNotInstalled = true;
                    this.UpdateTextLabel(this.statut, this.ressourceManager.GetString(ResourceMessage.STATUS_ERROR_OFFICE));
                    MessageBox.Show(this.ressourceManager.GetString(ResourceMessage.STATUS_ERROR_OFFICE));
                }
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
                this.ProcessNbFiles(openFileDialog1.FileNames);
                this.progressBar1.Value = 1;
                foreach (string file in openFileDialog1.FileNames)
                {
                    // Create a PictureBox.
                    if (Path.GetExtension(file) != FileExtension.PDF)
                    {
                        this.Converts(file);
                    }
                }
            }

            this.InSleep();
        }

        /// <summary>
        /// Updates all the UI colors to make them appear as processing
        /// </summary>
        private void InAction()
        {
            this.ChangeMode(Color.FromArgb(197, 56, 39), Color.White, PdfConverter.Properties.Resources.drop);            
        }

        /// <summary>
        /// Updates all the UI colors to make them appear as not processing
        /// </summary>
        private void InSleep()
        {
            this.ChangeMode(Color.White, Color.FromArgb(197, 56, 39), PdfConverter.Properties.Resources.drag);
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
