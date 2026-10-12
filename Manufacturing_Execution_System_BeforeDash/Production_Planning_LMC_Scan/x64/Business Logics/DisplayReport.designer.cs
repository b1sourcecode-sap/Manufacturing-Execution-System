using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
namespace InvoicePrinting_Sujal
{
    partial class DisplayReport : System.Windows.Forms.Form
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public DisplayReport()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}


        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DisplayReport));
            this.crystalReportViewer1 = new CrystalDecisions.Windows.Forms.CrystalReportViewer();
            this.ExportExcel = new System.Windows.Forms.Button();
            this.ExportPDF = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // crystalReportViewer1
            // 
            this.crystalReportViewer1.ActiveViewIndex = -1;
            this.crystalReportViewer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.crystalReportViewer1.Cursor = System.Windows.Forms.Cursors.Default;
            this.crystalReportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.crystalReportViewer1.Location = new System.Drawing.Point(0, 0);
            this.crystalReportViewer1.Name = "crystalReportViewer1";
            this.crystalReportViewer1.SelectionFormula = "";
            this.crystalReportViewer1.ShowGroupTreeButton = false;
            this.crystalReportViewer1.Size = new System.Drawing.Size(943, 582);
            this.crystalReportViewer1.TabIndex = 0;
            this.crystalReportViewer1.ToolPanelView = CrystalDecisions.Windows.Forms.ToolPanelViewType.None;
            this.crystalReportViewer1.ToolPanelWidth = 0;
            this.crystalReportViewer1.ViewTimeSelectionFormula = "";
            // 
            // ExportExcel
            // 
            this.ExportExcel.Image = ((System.Drawing.Image)(resources.GetObject("ExportExcel.Image")));
            this.ExportExcel.Location = new System.Drawing.Point(485, 1);
            this.ExportExcel.Name = "ExportExcel";
            this.ExportExcel.Size = new System.Drawing.Size(29, 30);
            this.ExportExcel.TabIndex = 2;
            this.ExportExcel.UseVisualStyleBackColor = true;
            this.ExportExcel.Click += new System.EventHandler(this.ExportExcel_Click);
            // 
            // ExportPDF
            // 
            this.ExportPDF.Image = ((System.Drawing.Image)(resources.GetObject("ExportPDF.Image")));
            this.ExportPDF.Location = new System.Drawing.Point(450, 1);
            this.ExportPDF.Name = "ExportPDF";
            this.ExportPDF.Size = new System.Drawing.Size(29, 30);
            this.ExportPDF.TabIndex = 3;
            this.ExportPDF.UseVisualStyleBackColor = true;
            this.ExportPDF.Click += new System.EventHandler(this.ExportPDF_Click);
            // 
            // DisplayReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(943, 582);
            this.Controls.Add(this.ExportPDF);
            this.Controls.Add(this.ExportExcel);
            this.Controls.Add(this.crystalReportViewer1);
            this.Name = "DisplayReport";
            this.Text = "DisplayReport";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DisplayReport_FormClosing);
            this.Load += new System.EventHandler(this.DisplayReport_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private CrystalDecisions.Windows.Forms.CrystalReportViewer crystalReportViewer1;
        private System.Windows.Forms.Button ExportExcel;
        private System.Windows.Forms.Button ExportPDF;

        private void rpt_Load(object sender, System.EventArgs e)
        {

        }
    }
}