using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace InvoicePrinting_Sujal
{

    public partial class DisplayReport : Form
    {
        InvoicePrinting_Sujal.Class.GetFileNameClass oGetFileName;
        string FileType = "";

        private string RptName;

        public string ReportName
        {
            get { return RptName; }
            set { RptName = value; }
        }
	
        class Test
        {
            SaveFileDialog _oFileDialog;
           

            public Test()
            {
                _oFileDialog = new SaveFileDialog();
            }
            public void openDial(Object ptr)
            {
                WindowWrapper oWindow = new WindowWrapper((IntPtr)ptr);
                if (_oFileDialog.ShowDialog(oWindow) != DialogResult.OK)
                {
                    _oFileDialog.FileName = string.Empty;
                }
                oWindow = null;
            }
        }

        CrystalDecisions.CrystalReports.Engine.ReportDocument CrRptDI = null;
        DataSet ds = new DataSet();

        public DisplayReport(DataSet dsReport)
        {
            InitializeComponent();
            ds = dsReport;
        }
        
        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {
            private IntPtr _hwnd;

            // Property
            public virtual IntPtr Handle
            {
                get { return _hwnd; }
            }

            // Constructor
            public WindowWrapper(IntPtr handle)
            {
                _hwnd = handle;
            }
        }
        //End Class For File Dialog Box
  
        private void DisplayReport_Load(object sender, EventArgs e)
        {
            displayReport();
            this.BringToFront();
        }
        void displayReport()
        {
            CrRptDI = new CrystalDecisions.CrystalReports.Engine.ReportDocument();
            CrRptDI.Load(Application.StartupPath + "\\" + "Reports" + "\\" + ReportName);

            //CrRptDI.Load(Application.StartupPath + "\\" + "Reports" + "\\" + "Net Position Report.rpt");
            
            CrRptDI.Database.Tables[0].SetDataSource(ds.Tables[0]);
            crystalReportViewer1.ReportSource = CrRptDI;

            crystalReportViewer1.ShowGroupTreeButton = false;
            crystalReportViewer1.DisplayGroupTree = false;
            crystalReportViewer1.ShowExportButton = false;
        }

        
        public void OpenDialog(string Fileext)
        {
            oGetFileName = new InvoicePrinting_Sujal.Class.GetFileNameClass();
            FileType = Fileext;
            if (Fileext == "Pdf")
            {
                oGetFileName.Filter = ".Pdf Files (*.Pdf)|*.Pdf";
            }
            else if (Fileext == "Xls")
            {
                oGetFileName.Filter = ".Excel Files (*.Xls)|*.Xls";
            }
            oGetFileName.EXportDone +=new EventHandler(oGetFileName_EXportDone);
            oGetFileName.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            Thread newThread = new Thread(new ThreadStart(oGetFileName.GetFileName));
            newThread.ApartmentState = ApartmentState.STA;
            
            try
            {
                newThread.Start();
                
                while (!newThread.IsAlive) ; // Wait for thread to get started
                Thread.Sleep(1);  // Wait a sec more
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            newThread = null;
        }

        private void oGetFileName_EXportDone(object sender, EventArgs e)
        {
            // Use file name as you will here
            InvoicePrinting_Sujal.Class.GetFileNameClass strValue = (InvoicePrinting_Sujal.Class.GetFileNameClass)sender;

            string val = strValue.FileName;
            try
            {
                if (FileType == "Pdf")
                {
                    CrRptDI.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, val);
                }
                else if (FileType == "Xls")
                {
                    CrRptDI.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.ExcelRecord, val);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void delay()
        {
            Thread.Sleep(1000);
        }
        private void ExportExcel_Click(object sender, EventArgs e)
        {
            try
            {
                OpenDialog("Xls");
                oGetFileName = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ExportPDF_Click(object sender, EventArgs e)
        {
            try
            {
                OpenDialog("Pdf");
                oGetFileName = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void DisplayReport_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (CrRptDI != null)
            {
                CrRptDI.Close();
                CrRptDI.Dispose();
            }
        }

    }
}