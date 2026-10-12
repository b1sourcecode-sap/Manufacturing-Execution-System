using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace InvoicePrinting_Sujal.Class
{
    class GetFileNameClass
    {

        #region DECLARE VARIABLE
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        SaveFileDialog _oFileDialog;
        FolderBrowserDialog _oFolderDialog;

        #endregion

        #region PROPERTIES

        public string FileName
        {
            get
               { return _oFileDialog.FileName; }
            set { _oFileDialog.FileName = value; }
        }

        public string Filter
        {
            get { return _oFileDialog.Filter; }
            set { _oFileDialog.Filter = value; }
        }

        public string InitialDirectory
        {
            get { return _oFileDialog.InitialDirectory; }
            set { _oFileDialog.InitialDirectory = value; }
        }
        public string GetFolderPath
        {
            get { return _oFolderDialog.SelectedPath; }
            set { _oFolderDialog.SelectedPath = value; }
        }

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public GetFileNameClass()
        {
            _oFolderDialog = new FolderBrowserDialog();
            _oFileDialog = new SaveFileDialog();
        }

        #endregion

        #region CLASS METHODS

        public void GetFileName()
        {
            IntPtr ptr = GetForegroundWindow();
            WindowWrapper oWindow = new WindowWrapper(ptr);
            DialogResult dr = _oFileDialog.ShowDialog(oWindow);
            if (dr != DialogResult.OK)
            {
                _oFileDialog.FileName = string.Empty;
            }
            EventArgs e = new EventArgs();
            EXportDone (this,e); 
            oWindow = null;
        } // End of GetFileName
        public void GetFolderName()
        {
            IntPtr ptr = GetForegroundWindow();
            WindowWrapper oWindow = new WindowWrapper(ptr);
            if (_oFolderDialog.ShowDialog(oWindow) != DialogResult.OK)
            {
                _oFolderDialog.SelectedPath = string.Empty;
            }
            oWindow = null;
        }

        public event EventHandler EXportDone; // End of GetFolderName

        #endregion
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
}

