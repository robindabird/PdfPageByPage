// <copyright file="PdfConvert.Designer.cs" company="OpenSource">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Robin Portigliatti</author>
// <date>01/05/2017 </date>
namespace PdfConverter
{
    /// <summary>
    /// Class representing the form of the PDFCONVERT application
    /// </summary>
    public partial class PdfConvert
    {
        /// <summary>
        /// Variable needed for the designer.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean all used resources.
        /// </summary>
        /// <param name="disposing">true if the managed resources has to be disposed, false otherwise.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }

        #region Code généré par le Concepteur Windows Form

        /// <summary>
        /// Méthode requise pour la prise en charge du concepteur - ne modifiez pas
        /// le contenu de cette méthode avec l'éditeur de code.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PdfConvert));
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.document = new System.Windows.Forms.Label();
            this.statut = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.Window;
            this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pictureBox1.Image = global::PdfConverter.Properties.Resources.drag;
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(494, 341);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.PictureBox1_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.WindowText;
            this.label1.Location = new System.Drawing.Point(436, 315);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "0/0";
            // 
            // document
            // 
            this.document.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.document.AutoSize = true;
            this.document.ForeColor = System.Drawing.SystemColors.WindowText;
            this.document.Location = new System.Drawing.Point(12, 315);
            this.document.MaximumSize = new System.Drawing.Size(114, 17);
            this.document.MinimumSize = new System.Drawing.Size(114, 17);
            this.document.Name = "document";
            this.document.Size = new System.Drawing.Size(114, 17);
            this.document.TabIndex = 3;
            this.document.Text = "Aucun document";
            // 
            // statut
            // 
            this.statut.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.statut.AutoSize = true;
            this.statut.ForeColor = System.Drawing.SystemColors.WindowText;
            this.statut.Location = new System.Drawing.Point(154, 314);
            this.statut.MinimumSize = new System.Drawing.Size(174, 17);
            this.statut.Name = "statut";
            this.statut.Size = new System.Drawing.Size(174, 17);
            this.statut.TabIndex = 4;
            this.statut.Text = "Aucun traitement en cours";
            // 
            // progressBar1
            // 
            this.progressBar1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBar1.Location = new System.Drawing.Point(15, 288);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(467, 23);
            this.progressBar1.TabIndex = 5;
            // 
            // PdfConvert
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(494, 341);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.statut);
            this.Controls.Add(this.document);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(512, 388);
            this.MinimumSize = new System.Drawing.Size(512, 388);
            this.Name = "PdfConvert";
            this.Text = "HCD - PdfSplitter";
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label document;
        private System.Windows.Forms.Label statut;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}