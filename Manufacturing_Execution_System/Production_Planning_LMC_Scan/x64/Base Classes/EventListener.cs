using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json;
using SAPbobsCOM;
using SAPbouiCOM;


namespace Production_Planning_LMC
{
    /// <summary>
    /// public sealed class EventListener
    /// </summary>
    public sealed class EventListener
    {
        #region FORM VARIABLES DECLARATION
        private static EventListener    _Listener;
        private SAPbouiCOM.Application  _Application;
        private SAPbobsCOM.Company      _Company;
        private Base                    _Object;
        private Hashtable               _Collection;
        private Hashtable               _LookCollection;
        private string                  _FormUID;
        private SerialPort _scannerPort;
        string FirstScan = string.Empty;
        string SecondScan = string.Empty;
        int scancount = 0;
        private SAPbobsCOM.Recordset _RSWorkId, _RSRoute, _RSWorkExecDetails, _RSWorkOrderDetails, _RSD, _RSBom, _RSPP;
        string route = string.Empty;
        string routeId = string.Empty;
        string WorkId = string.Empty;
        DateTime LDurationTime;
        private bool _isSystemClear = false;
        string productionOrderEntry = string.Empty;
        public SAPbouiCOM.Matrix _Matrix;
        public SAPbouiCOM.Form oForm;
        private static SAPbouiCOM.Form oScanForm;
        private static SAPbouiCOM.EditText oScanTxt;
        private string _scanBuffer = "";
        private string routeDsp, productName, model;
        private string _MESDBTime = "";
        System.Windows.Forms.Timer focusTimer = new System.Windows.Forms.Timer();
        private SAPbouiCOM.Form _oForm;
        double totalBreakMinutes;
        string totalBrkDurationString;
        // Add this at the top of your class to prevent opening multiple browser tabs
        private static bool _isDashboardOpen = false;
        private static System.Windows.Forms.Timer _refreshTimer = null;
        private static string _slSessionId = "";
        private static string _slUrl = "https://192.168.11.200:50000/b1s/v1/";
        private static string _slCookieHeader = "";
        private bool _isProcessingIssue = false;
        #endregion

        #region CONSTRUCTOR

        private EventListener()
        {
            this.Connect();

            this._Application.MenuEvent       += new _IApplicationEvents_MenuEventEventHandler(SboApplication_MenuEvent);
            this._Application.ItemEvent       += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SboApplication_ItemEvent);
            this._Application.AppEvent        += new _IApplicationEvents_AppEventEventHandler(SboApplication_AppEvent);
            this._Application.RightClickEvent += new _IApplicationEvents_RightClickEventEventHandler(SboApplication_RightClickEvent);
            this._Application.FormDataEvent   += new _IApplicationEvents_FormDataEventEventHandler(SboApplication_FormDataEvent);
            
            

            this._Collection                      = new Hashtable(10, (float)0.5);
            this._LookCollection                  = new Hashtable(10, (float)0.5);
        }

        public static EventListener getEventListener()
        {
            if (_Listener == null)
                _Listener = new EventListener();


            return _Listener;
        }

        #endregion

        #region PROPERTIES

        public SAPbobsCOM.Company Company
        {
            get { return this._Company; }
        }

        public SAPbouiCOM.Application SBO_Application
        {
            get { return this._Application; }
        }

        public Hashtable Collection
        {
            get { return this._Collection; }
        }

        public Hashtable LookUpCollection
        {
            get { return this._LookCollection; }
        }

        #endregion

        #region COMPANY CONNECT

        private void Connect()
        {
            this.SetApplication();
            this.ConnectToCompany();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi;
            string               oConnectionString;
            try
            {
                if (Environment.GetCommandLineArgs().Length > 0)
                {
                    //oSboGuiApi = new SAPbouiCOM.SboGuiApiClass();
                    oSboGuiApi = new SAPbouiCOM.SboGuiApi();
                    oConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                    oSboGuiApi.Connect(oConnectionString);
                    this._Application = oSboGuiApi.GetApplication(-1);
                    //this._Company = new SAPbobsCOM.CompanyClass();
                    this._Company = new SAPbobsCOM.Company();
                }
                else
                {
                    throw new Exception("Connection string missing.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private void ConnectToCompany()
        {
            string oCookie, oConnectionContext;
            try
            {
                oCookie = this._Company.GetContextCookie();
                oConnectionContext = this._Application.Company.GetConnectionContext(oCookie);
                this._Company.SetSboLoginContext(oConnectionContext);

                if (this._Company.Connect() != 0)
                {
                    _Application.StatusBar.SetText("Add-On is not Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception(_Company.GetLastErrorDescription());
                }
                else
                {
                    _Application.StatusBar.SetText("Add-On is Connecting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    //OpenUDOScanForm();
                    CreateHiddenListenerForm();
                    StartFocusTimer(oForm);
                    ForceScanFocus();
                    //OpenTrimAssemblyStatusBoard(false);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region EVENT HANDLERS

        #region MENU EVENT

        private void SboApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                {

                    switch (pVal.MenuUID)
                    {
                        //case Constants.User_Menus.MENU_LGProduction:
                        //     this._Object = new ClsLGProductionOrder();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ProductionPlanning);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsLGProductionOrder)_Object).FormDefault();
                        //     break;
                        // case Constants.User_Menus.MENU_LGApproval:
                        //     this._Object = new ClsApprovalTemplate();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovalTemplate);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsApprovalTemplate)_Object).FormDefault();
                        //     break;
                        // case Constants.User_Menus.MENU_ABPC:
                        //     this._Object = new ClsApprovedBPCreation();
                        //     Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovedBPCreation);
                        //     this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //     ((ClsApprovedBPCreation)_Object).FormDefault();
                        //     break;
                        //case Constants.User_Menus.MENU_ABPCApproval:
                        //    this._Object = new ClsApprovalTemplate();
                        //    Utilities.LoadForm(ref this._Object, Constants.Forms.ABPTemplate);
                        //    this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //    ((ClsApprovalTemplate)_Object).FormDefault();
                        //    break;
                        case Constants.User_Menus.MENU_EngineChasis:
                            this._Object = new ProductionPlanningLMC();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.EngineChasisMapping);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ProductionPlanningLMC)_Object).FormDefault();
                            break;
                        case Constants.User_Menus.MENU_WorkOrderDetailss:
                            this._Object = new ClsWorkOrderDetails();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.WorkOrderDetailss);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsWorkOrderDetails)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_JobOrderExecutions:
                            this._Object = new ClsJobOrderExecution();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsJobOrderExecution)_Object).FormDefault();
                            break;
                        case Constants.User_Menus.MENU_LOTMASTER:
                            Utilities.Application.SBO_Application.Menus.Item("LOTNOMASTER").Activate();
                            break;
                        case Constants.User_Menus.MENU_OCNMASTER:
                            Utilities.Application.SBO_Application.Menus.Item("OCNMASTER1").Activate();
                            break;
                        case Constants.User_Menus.MENU_PRODUCTIONORDER:
                            Utilities.Application.SBO_Application.Menus.Item("4369").Activate();
                            break;
                        case Constants.User_Menus.MENU_INVENTORYTRANSFERREQ:
                            Utilities.Application.SBO_Application.Menus.Item("3088").Activate();
                            break;
                        case Constants.User_Menus.MENU_INVENTORYTRANSFER:
                            Utilities.Application.SBO_Application.Menus.Item("3080").Activate();
                            break;
                        case Constants.User_Menus.MENU_ISSUEPROD:
                            Utilities.Application.SBO_Application.Menus.Item("4371").Activate();
                            break;
                        case Constants.User_Menus.MENU_RECEIPTPROD:
                            Utilities.Application.SBO_Application.Menus.Item("4370").Activate();
                            break;
                        case Constants.User_Menus.MENU_DASHBOARD:
                            OpenTrimAssemblyStatusBoard(true);
                            break;
                        case Constants.User_Menus.MENU_PRODUCTIONMASTER:
                            this._Object = new ClsProductionMaster();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.ProductionMaster);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsProductionMaster)_Object).FormDefault();
                            break;

                        // System Menu's
                        case Constants.System_Menus.mnu_ADD:
                        case Constants.System_Menus.mnu_FIND:
                        case Constants.System_Menus.mnu_FIRST:
                        case Constants.System_Menus.mnu_LAST:
                        case Constants.System_Menus.mnu_NEXT:
                        case Constants.System_Menus.mnu_PREVIOUS:
                        case Constants.System_Menus.mnu_REFRESH:
                        case Constants.System_Menus.mnu_ADD_ROW:
                        case Constants.System_Menus.mnu_DELETE_ROW:
                        case Constants.System_Menus.mnu_SALES_ORDER:
                        case Constants.System_Menus.mnu_OUTGOING_PAYMENTS:
                        case Constants.System_Menus.mnu_GL_ACCOUNT_DETERMINATION:
                        case Constants.System_Menus.mnu_Duplicate:

                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_GROSS_PROFIT:
                        //    BubbleEvent = false;
                        //    break;
                        case Constants.System_Menus.mnu_ROW_DETAILS:
                            BubbleEvent = false;
                            break;

                        case Constants.System_Menus.mnu_DELETE_ROW:
                        case Constants.System_Menus.mnu_Duplicate:
                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region ITEM EVENT

        private void SboApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this._FormUID = FormUID;
               
                //Utilities.HideSystemMessege(ref pVal);
                if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) //
                {
                    switch (pVal.FormType)
                    {
                        //case Constants.System_Forms.AR_INVOICE:
                        //    this._Object = new ClsARInvoice();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsARInvoice)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsARInvoice)this._Object).FormDefault();
                        //    break;

                        //case Constants.System_Forms.AR_CREDITNOTE:
                        //    this._Object = new ClsARCreditNote();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsARCreditNote)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    break;

                        //case Constants.System_Forms.GOODS_RETURN:
                        //    this._Object = new ClsGoodsReturn();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsGoodsReturn)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsGoodsReturn)this._Object).FormDefault();
                        //    break;

                        //case Constants.System_Forms.Production_ORDER:
                        //    this._Object = new ClsProductionOrder();
                        //    _Collection.Add(FormUID, this._Object);
                        //    ((ClsProductionOrder)this._Object).Form = _Application.Forms.Item(FormUID);
                        //    ((ClsProductionOrder)this._Object).FormDefault();
                        //    break;

                        case Constants.System_Forms.Production_ORDER:
                            this._Object = new ClsProductionOrderr();
                            _Collection.Add(FormUID, this._Object);
                            ((ClsProductionOrderr)this._Object).Form = _Application.Forms.Item(FormUID);
                            ((ClsProductionOrderr)this._Object).FormDefault();
                            break;
                    }
                }

                if (this._Collection.Contains(FormUID))
                {
                    this._Object = (Production_Planning_LMC.Base)this._Collection[FormUID];

                    if (pVal.BeforeAction && this._Object.IsLookUpOpen)
                    {
                        BubbleEvent = false;
                        _Application.Forms.Item(this._Object.LookUpUID).Select();
                    }
                    else
                    {
                        this._Object.Item_Event(FormUID, ref pVal, ref BubbleEvent);
                    }

                    if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    {
                        if (_LookCollection.ContainsKey(FormUID))
                        {
                            this._Object = (Production_Planning_LMC.Base)_Collection[_LookCollection[FormUID]];
                            _LookCollection.Remove(FormUID);
                            this._Object.IsLookUpOpen = false;
                        }
                        this._Object.Dispose();
                        _Collection.Remove(FormUID);
                        System.GC.Collect();
                    }
                }

                if (FormUID == "frmBarScan" && pVal.ItemUID == "txtScan" && !pVal.BeforeAction)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_LOST_FOCUS)
                    {
                        ForceScanFocus(); 
                    }
                }

                /////////////////////////
                if (oScanTxt == null) return;

                if (FormUID == "frmBarScan" && pVal.ItemUID == "txtScan" && !pVal.BeforeAction)
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN)
                    {
                        int key = pVal.CharPressed;

                        // 2️⃣ Ignore System Keys (Shift, Ctrl, etc which are -1 or 0)
                        if (key <= 0) return;

                        // 3️⃣ Check for ENTER (13)
                        if (key == 13 || key == 10)
                        {
                            try
                            {
                                // ✅ Process the accumulated buffer, NOT oScanTxt.Value
                                string finalBarcode = _scanBuffer.Trim();

                                if (!string.IsNullOrEmpty(finalBarcode))
                                {
                                    // --- YOUR LOGIC HERE ---
                                    Utilities.Application.SBO_Application.StatusBar.SetText(
                                        "Processing: " + finalBarcode,
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                                    // Clear Buffer & TextBox after success
                                    // _scanBuffer = "";
                                    try { oScanTxt.Value = ""; } catch { }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("Error: " + ex.Message);
                                //  _scanBuffer = ""; // Reset on error
                            }

                            // Focus ready for next scan
                            ForceScanFocus();
                        }
                        // 4️⃣ Check for Backspace (8) - Optional but good for manual correction
                        else if (key == 8)
                        {
                            if (_scanBuffer.Length > 0)
                                _scanBuffer = _scanBuffer.Substring(0, _scanBuffer.Length - 1);
                        }
                        // 5️⃣ Accumulate Characters (If not Enter)
                        else
                        {
                            // Convert ASCII code to Character and add to buffer
                            _scanBuffer += (char)key;

                        }
                    }
                }
                //////////////////////

                #region Scanning Chassis No and its respective Actions
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction == false && pVal.InnerEvent == true)
                {
                    SAPbouiCOM.Form _Form = Utilities.Application.SBO_Application.Forms.ActiveForm;
                    if (pVal.FormTypeEx == "SCAN_BAR" && pVal.ItemUID == "txtScan" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction == false)
                    {

                        SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)_Form.Items.Item("txtScan").Specific;
                        //SAPbouiCOM.Item addBtn = _Form.Items.Item("1"); // The Add Button (UI ID 1)

                        string scannedValue = _scanBuffer;

                        if (string.IsNullOrEmpty(scannedValue))
                        {
                            return;
                        }

                        // --- FIRST SCAN ---
                        if (scancount == 0)
                        {
                            // 1. Process the First Scan String
                            //string first = scannedValue.ToUpper();
                            //if (first.Length > 1)
                            //{
                            //    first = first.Substring(0, first.Length - 1) + "*" + first.Substring(first.Length - 1);
                            //}

                            string first = scannedValue.ToUpper();

                            if (first.Length > 1)
                            {
                                int secondLastIndex = first.Length - 2;

                                // Case 1: second last character is '8' → replace with '*'
                                if (first.Length == 19)
                                {
                                    if (first[secondLastIndex] == '8')
                                    {
                                        first = first.Substring(0, secondLastIndex) + "*" + first.Substring(secondLastIndex + 1);
                                    }
                                }
                                else if (first.Length == 18) {
                                // Case 2: no '*' exists → insert '*' before last character
                                if (!first.Contains("*"))
                                    {
                                        first = first.Substring(0, first.Length - 1) + "*" + first.Substring(first.Length - 1);
                                    }
                                }

                                int checkIndex = first.Length - 2;

                                if (checkIndex < 17)
                                {
                                    //if (first.Length == 17 && !first.StartsWith("M"))
                                    //{
                                    //    first = "M" + first;
                                    //}
                                    // 1. Ping the server
                                    try
                                    {
                                        using (System.Net.NetworkInformation.Ping pingSender = new System.Net.NetworkInformation.Ping())
                                        {
                                            pingSender.Send("192.168.11.200", 500);
                                        }
                                    }
                                    catch { /* Ignore ping errors */ }

                                    //// 2. Prepare data for the browser
                                    //string currentTime = DateTime.Now.ToString("HH:mm:ss");

                                    //// 3. Show the Browser Message
                                    //ShowBrowserMessage(
                                    //    icon: "❌",
                                    //    subTitle: "INVALID CHASSIS",
                                    //    chassis: scannedValue, model: "N/A", modelName: "N/A", line: "", time: currentTime,color: "#dc3545", nextAction: "Invalid chassis. Please scan again.");

                                    //// 4. Reset buffer and stop SAP event
                                    //_scanBuffer = "";
                                    //BubbleEvent = false;
                                    //return;
                                }
                            }
                            SAPbobsCOM.Recordset rsTime = null;
                            try
                            {
                                rsTime = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsTime.DoQuery("SELECT \"U_MESBDTime\" FROM \"OADM\"");
                                if (!rsTime.EoF)
                                {

                                    _MESDBTime = rsTime.Fields.Item("U_MESBDTime").Value;
                                }
                            }
                            catch (Exception ex)
                            {
                            }
                            finally
                            {
                                if (rsTime != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsTime);
                                    rsTime = null;
                                }

                            }

                            FirstScan = first;
                            scancount = 1;
                            OpenScanInBrowser(first);
                            // 2. Clear the field visually so it is ready for Second Scan
                            _scanBuffer = "";

                            // IMPORTANT: Stop the event here. 
                            // This keeps the focus in Item "14" and prevents SAP from checking OITM table.
                            BubbleEvent = false;
                            ForceScanFocus();

                            // Show a small status message so user knows Scan 1 was accepted
                            // SBO_Application.MessageBox("First scan accepted. Please scan secondary code...");
                            //CloseSystemPopupNo();

                        }
                        // --- SECOND SCAN ---
                        else if (scancount == 1)
                        {
                            //
                            SecondScan = scannedValue.ToUpper().Replace(" ", "_");
                            string matchedColumn = "";
                                //SecondScan = scannedValue.ToUpper();
                                SAPbobsCOM.Recordset _RSWorkId = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string query = $@"SELECT ""U_FirstBarcode"", ""U_LastBarcode"", ""U_JobPause"", ""U_JobResume""  
                                             FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{FirstScan}'";
                            _RSWorkId.DoQuery(query);
                                if (_RSWorkId.RecordCount > 0)
                                {
                                    while (!_RSWorkId.EoF)  // loop rows
                                    {
                                        for (int i = 0; i < _RSWorkId.Fields.Count; i++)
                                        {
                                            var field = _RSWorkId.Fields.Item(i);
                                            string fieldValue = field.Value != null ? field.Value.ToString().Trim() : "";

                                            if (!string.IsNullOrEmpty(fieldValue) && fieldValue.Equals(SecondScan, StringComparison.OrdinalIgnoreCase))
                                            {
                                                matchedColumn = field.Name;
                                                break;
                                            }

                                        }
                                        if (!string.IsNullOrEmpty(matchedColumn))
                                            break; // stop searching if match is found

                                        _RSWorkId.MoveNext(); // go to next row
                                    }

                                if (string.IsNullOrEmpty(matchedColumn))
                                {
                                    ShowBrowserMessage("❌", "INVALID ACTION", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Barcode not recognized for this chassis. Scan valid action.");
                                    //SBO_Application.MessageBox("Invalid Barcode! This code is not associated with the scanned Chassis actions.");
                                    _scanBuffer = "";
                                    scancount = 0; // Reset to start
                                    BubbleEvent = false;
                                    return;
                                }


                                SAPbobsCOM.Recordset _RSRoute = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query_route = $@"SELECT T0.""LineId"",T0.""U_RouteDsp"",T0.""U_RouteId"", T0.""U_WorkId"", T1.""U_PrdOrdEnt"", T0.""U_RouteIdNum"",T1.""U_ProductName"",T1.""U_Model"", T0.""U_JobExeNo""
                                                     FROM ""@WRKORDRDTLSC"" T0 INNER JOIN ""@WRKORDRDTLSH"" T1 on T0.""DocEntry"" = T1.""DocEntry"" 
                                                     WHERE T0.""U_ChasisNo"" = '{FirstScan}' and ""{matchedColumn}""='{SecondScan}'";
                                _RSRoute.DoQuery(query_route);

                                if (_RSRoute.RecordCount > 0)
                                {
                                   
                                     route = _RSRoute.Fields.Item("U_RouteDsp").Value;
                                    //routeId = _RSRoute.Fields.Item("U_RouteIdNum").Value;
                                    //routeId = _RSRoute.Fields.Item("U_RouteId").Value;
                                    string rawRouteId = _RSRoute.Fields.Item("U_RouteId").Value.ToString();

                                    // Check if it's not empty to avoid errors
                                    if (!string.IsNullOrEmpty(rawRouteId))
                                    {
                                        // Split by '-' and take the first item (Index 0), then Trim whitespace
                                        // Result: "1"
                                        routeId = rawRouteId.Split('-')[0].Trim();
                                    }
                                    else
                                    {
                                        routeId = "";
                                    }
                                    WorkId = _RSRoute.Fields.Item("U_WorkId").Value;
                                    productionOrderEntry = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString();
                                     routeDsp = _RSRoute.Fields.Item("U_RouteDsp").Value.ToString();
                                    productName = _RSRoute.Fields.Item("U_ProductName").Value.ToString();
                                    model = _RSRoute.Fields.Item("U_Model").Value.ToString();

                                    int currentLineId = Convert.ToInt32(_RSRoute.Fields.Item("U_RouteIdNum").Value);
                                    if (Convert.ToInt32(currentLineId) > 1)
                                    {
                                        int previousLineId = currentLineId - 1;

                                        SAPbobsCOM.Recordset _RSCheckPrev = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        // Query to check if the previous line has a value in U_IssueForProdEntry
                                        // We use IFNULL/NULLIF to handle both NULLs and Empty strings as Empty
                                        string query_prev = $@"
                                        SELECT IFNULL(NULLIF(T0.""U_IssueForProdEntry"", ''), '') AS ""PrevIssue"", ""U_RouteDsp""
                                        FROM ""@WRKORDRDTLSC"" T0
                                        INNER JOIN  ""@WRKORDRDTLSH"" T1 on T0.""DocEntry"" = T1.""DocEntry"" 
                                        WHERE T0.""U_ChasisNo"" = '{FirstScan}'
                                          AND T0.""U_RouteIdNum"" = {previousLineId} ";

                                        _RSCheckPrev.DoQuery(query_prev);


                                        //issue check baseEntry equals production docentry
                                        SAPbobsCOM.Recordset _RSCheckIssue = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        if (_RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim() != "" )
                                        {
                                           
                                            string issueBaseEntry = $@"
                                        SELECT ""BaseEntry""
                                        FROM IGE1
                                        WHERE ""DocEntry"" = '{_RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim()}'";
                                            _RSCheckIssue.DoQuery(issueBaseEntry);
                                        }

                                        if (_RSCheckPrev.RecordCount > 0)
                                        {
                                            string prevIssueEntry = _RSCheckPrev.Fields.Item("PrevIssue").Value.ToString().Trim();

                                            // 3. Block logic if previous entry is empty
                                            if (string.IsNullOrEmpty(prevIssueEntry))
                                            {
                                                ShowBrowserMessage("⚠️", "SEQUENCE ERROR", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", $"ACTION: Previous step [{_RSCheckPrev.Fields.Item("U_RouteDsp").Value}] must be completed first.");
                                               // SBO_Application.MessageBox(
                                                //    $"Sequence Error: The previous operation ({_RSCheckPrev.Fields.Item("U_RouteDsp").Value}) is not completed yet.");

                                                _scanBuffer = "";
                                                scancount = 0;
                                                BubbleEvent = false;

                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSCheckPrev);
                                                return;
                                            }
                                        }
                                        if(_RSCheckIssue.RecordCount > 0)
                                        {
                                            string BaseEntry = _RSCheckIssue.Fields.Item("BaseEntry").Value.ToString().Trim();
                                            string prodDocEntry = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString().Trim();
                                            if (BaseEntry != prodDocEntry)
                                            {
                                                ShowBrowserMessage("❌", "INVALID ISSUE FOR PRODUCTION", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Invalid Issue for Production.");
                                               

                                                _scanBuffer = "";
                                                scancount = 0;
                                                BubbleEvent = false;

                                                System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSCheckIssue);
                                                return;
                                            }
                                        }
                                    }
                                    //OpenScanInBrowser(SecondScan,FirstScan,routeDsp,productName,model);
                                    //OpenTrimAssemblyStatusBoard();
                                }
                                string query_Execution = $@"SELECT ""DocEntry"",""U_JobID"",""U_JobDesc"",""U_ProdOrdNo"",""U_WorkOrdNo"",
                                                   ""U_FGCode"",""U_FGDesc"",""U_EngineNo"",""U_ChassisNo"",""U_BatchNo"",
                                                   ""U_Qty"",""U_Status"",""U_StartDate"",""U_StartTime"",""U_EndDate"",
                                                   ""U_EndTime"",""U_TotBrkTime"",""U_TotActTime"",""U_Operator""
                                            FROM ""@WJOBEXEH""
                                            WHERE ""U_JobID"" = '{WorkId}'";

                                SAPbobsCOM.Recordset _RSWorkExecDetails = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                _RSWorkExecDetails.DoQuery(query_Execution);
                                if (_RSWorkExecDetails.RecordCount > 0)
                                {
                                    string currentStatus = _RSWorkExecDetails.Fields.Item("U_Status").Value.ToString();

                                    // --- Load UDO Service ---
                                    SAPbobsCOM.CompanyService oCompanyService = Utilities.Application.Company.GetCompanyService();
                                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("WJOBEXE");
                                    SAPbobsCOM.GeneralDataParams oParams =
                                        (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                                            SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                                    oParams.SetProperty("DocEntry", _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString());
                                    SAPbobsCOM.GeneralData oGeneralData = oGeneralService.GetByParams(oParams);

                                    SAPbobsCOM.GeneralDataCollection oChildren = oGeneralData.Child("WJOBEXEC");
                                    SAPbobsCOM.GeneralData oChild;

                                    DateTime now = DateTime.Now;

                                    #region When current status is "Start" (Can add a pause or Stop)
                                    if (currentStatus == "Start")
                                    {
                                        if (matchedColumn == "U_JobPause")
                                        {
                                            oGeneralData.SetProperty("U_Status", "Pause");
                                            // Create Pause Row
                                            oChild = oChildren.Add();
                                            oChild.SetProperty("U_BrkStartDt", now);
                                            oChild.SetProperty("U_BrkStartTm", now);

                                            //SBO_Application.StatusBar.SetText("Job Paused",
                                            //    SAPbouiCOM.BoMessageTime.bmt_Short,
                                            //    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Paused!");
                                            //ShowBrowserMessage("⏸️", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", "ACTION: Job is suspended. Scan RESUME to continue.");
                                             ShowBrowserMessage("✅", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ffea00", "ACTION: Job is suspended. Scan RESUME to continue.");

                                            oGeneralService.Update(oGeneralData);
                                            this._Object = new ClsJobOrderExecution();
                                            Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);

                                            // 2. Get a reference to the SAP Form from your newly created object
                                            // Note: This assumes your ClsJobOrderExecution class has a public 'Form' or 'oForm' property
                                            SAPbouiCOM.Form oForm = ((ClsJobOrderExecution)this._Object).Form;

                                            if (oForm != null)
                                            {
                                                oForm.Freeze(true);
                                                try
                                                {
                                                    // 3. Switch the form to Find Mode
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                                    // 4. Set the DocEntry into your field (using the UID you mentioned: "txtDocE")
                                                    // If "txtDocE" is the UID on your custom .srf, this will work.
                                                    ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocE").Specific).Value = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();

                                                    // 5. Simulate clicking the "Find" button (UID "1")
                                                    oForm.Items.Item("1").Click();

                                                    SetRemarkFieldsState(oForm, "Pause");
                                                }
                                                catch (Exception ex)
                                                {
                                                    Utilities.Application.SBO_Application.StatusBar.SetText("Failed to find record: " + ex.Message);
                                                }
                                                finally
                                                {
                                                    oForm.Freeze(false);
                                                }
                                            }

                                        }

                                        //else if(matchedColumn == "U_LastBarcode")
                                        //{
                                        //    SAPbouiCOM.ProgressBar pb = null;
                                        //    try
                                        //    {
                                        //        pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);
                                        //        int goodIssue = Create_IssueFor_Production(_RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan, _RSRoute.Fields.Item("U_JobExeNo").Value.ToString());
                                        //        //int goodIssue = 790;

                                        //        if (goodIssue > 0)
                                        //        {
                                        //            //Only update Status and Times if Goods Issue succeeded
                                        //            oGeneralData.SetProperty("U_Status", "Stop");
                                        //            oGeneralData.SetProperty("U_EndDate", now);
                                        //            oGeneralData.SetProperty("U_EndTime", now);
                                        //            //oGeneralData.SetProperty("U_TotBrkTime", new DateTime(1899, 12, 30)); // 00:00:00 No breaks
                                        //            oGeneralData.SetProperty("U_TotBrkTime", "00:00");
                                        //            oGeneralData.SetProperty("U_IForProdEntry", goodIssue.ToString());

                                        //            //// Calculate Actual Time Consumed
                                        //            ////double totalActMinutes = 0;
                                        //            ////try
                                        //            ////{
                                        //            ////    object hStartDt = oGeneralData.GetProperty("U_StartDate");
                                        //            ////    object hStartTm = oGeneralData.GetProperty("U_StartTime");

                                        //            ////    object cEndDt = oGeneralData.GetProperty("U_EndDate");
                                        //            ////    object cEndTm = oGeneralData.GetProperty("U_EndTime");

                                        //            ////    DateTime dtHeaderStart = GetFullDateTime(hStartDt, hStartTm);
                                        //            ////    DateTime dtLastChildEnd = GetFullDateTime(cEndDt, cEndTm);

                                        //            ////    if (dtHeaderStart != DateTime.MinValue)
                                        //            ////        totalActMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                        //            ////}
                                        //            ////catch { }

                                        //            ////oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes < 0 ? 0 : totalActMinutes));
                                        //            double totalActMinutes = 0;
                                        //            string durationString = "";
                                        //            try
                                        //            {
                                        //                DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                        //                DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                        //                double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                        //                totalActMinutes = totalElapsedMinutes;

                                        //                int hourss = (int)(totalActMinutes / 60);
                                        //                int minss = (int)(totalActMinutes % 60);
                                        //                durationString = $"{hourss:D2}:{minss:D2}";
                                        //            }
                                        //            catch { }


                                        //            ////if alpha
                                        //            //oGeneralData.SetProperty("U_TotActTime", durationString);

                                        //            // NOTE: Visually in SAP, this will show "01:16", but the underlying value 
                                        //            // will contain the extra day (1899-12-31) which is important for the next step.
                                        //            //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes));


                                        //            // Update External Tables 
                                        //            try
                                        //            {

                                        //                DateTime dtStartDate = (DateTime)oGeneralData.GetProperty("U_StartDate");
                                        //                DateTime dtStartTime = (DateTime)oGeneralData.GetProperty("U_StartTime");
                                        //                DateTime dtEndDate = (DateTime)oGeneralData.GetProperty("U_EndDate");
                                        //                DateTime dtEndTime = (DateTime)oGeneralData.GetProperty("U_EndTime");
                                        //                //DateTime dtBrkTime = (DateTime)oGeneralData.GetProperty("U_TotBrkTime");
                                        //                //DateTime dtActTime = (DateTime)oGeneralData.GetProperty("U_TotActTime");
                                        //                string strBrkTime = Convert.ToString(oGeneralData.GetProperty("U_TotBrkTime"));
                                        //                string strActTime = Convert.ToString(oGeneralData.GetProperty("U_TotActTime"));

                                        //                string sqlStartDate = dtStartDate.ToString("yyyyMMdd");
                                        //                string sqlEndDate = dtEndDate.ToString("yyyyMMdd");
                                        //                string sqlStartTime = dtStartTime.ToString("HHmm");
                                        //                string sqlEndTime = dtEndTime.ToString("HHmm");
                                        //                //string sqlBrkTime = dtBrkTime.ToString("HHmm");
                                        //                //string sqlActTime = dtActTime.ToString("HHmm");
                                        //                string sqlBrkTime = strBrkTime.Replace(":", "").PadLeft(4, '0');
                                        //                string sqlActTime = strActTime.Replace(":", "").PadLeft(4, '0');


                                        //                /// for exceeding 24 hours 
                                        //                DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                        //                DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                        //                double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                        //                totalActMinutes = totalElapsedMinutes;

                                        //                int totalHrs = (int)(totalActMinutes / 60);
                                        //                int totalMns = (int)(totalActMinutes % 60);

                                        //                // Use a string format for the SQL update
                                        //                // This will result in "2516" instead of "0116"
                                        //                string manualTimeStr = $"{totalHrs:D2}:{totalMns:D2}";
                                        //                /// end for exceeding 24 hours


                                        //                SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        //                //for 24 hours exceed
                                        //                string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                        //                SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                        //                ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                        //                ""U_TotalBrkDwnT"" = '{strBrkTime}', ""U_ActualTimeCon"" = '{manualTimeStr}'
                                        //                WHERE ""U_WorkId"" = '{WorkId}'";
                                        //                rsUpdate.DoQuery(updateQuery);
                                        //                /////////

                                        //                string chasisNo = oGeneralData.GetProperty("U_ChassisNo");

                                        //                // Check if other processes are still pending
                                        //                SAPbobsCOM.Recordset rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        //                string selectQuery = $@"SELECT * FROM ""@WRKORDRDTLSC"" 
                                        //            WHERE ""U_ChasisNo"" = '{chasisNo}' and ""U_IssueForProdEntry"" IS NULL ";
                                        //                rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        //                rsSelect.DoQuery(selectQuery);

                                        //                if (rsSelect.RecordCount > 0)
                                        //                {
                                        //                    SBO_Application.StatusBar.SetText("Goods Issue Created, but Engine/Chassis status not completed (other processes pending).",
                                        //                        SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //                    //SBO_Application.MessageBox("Goods Issue Created, but Engine/Chassis status not completed (other processes pending).");
                                        //                }
                                        //                else
                                        //                {
                                        //                    try
                                        //                    {
                                        //                        int totalMinutes = 0;
                                        //                        string sumQuery = $@"SELECT ""U_ActualTimeCon"" FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}'";

                                        //                        rsSelect.DoQuery(sumQuery);

                                        //                        while (!rsSelect.EoF)
                                        //                        {
                                        //                            string timeStr = rsSelect.Fields.Item("U_ActualTimeCon").Value.ToString().Replace(":", "").Trim();

                                        //                            if (!string.IsNullOrEmpty(timeStr))
                                        //                            {
                                        //                                // Pad to ensure we have at least HHmm (e.g., "15" -> "0015", "130" -> "0130")
                                        //                                timeStr = timeStr.PadLeft(4, '0');

                                        //                                // Take the last 4 digits only to be safe
                                        //                                if (timeStr.Length > 4)
                                        //                                    timeStr = timeStr.Substring(timeStr.Length - 4);

                                        //                                if (int.TryParse(timeStr.Substring(0, 2), out int hours) &&
                                        //                                    int.TryParse(timeStr.Substring(2, 2), out int mins))
                                        //                                {
                                        //                                    totalMinutes += (hours * 60) + mins;
                                        //                                }
                                        //                            }
                                        //                            rsSelect.MoveNext();
                                        //                        }

                                        //                        // Convert total minutes back to HHmm format
                                        //                        int finalHours = totalMinutes / 60;
                                        //                        int finalMins = totalMinutes % 60;
                                        //                        string finalTotalTime = $"{finalHours:00}:{finalMins:00}";

                                        //                        //Update 
                                        //                        string updateEngineChasis = $@"UPDATE ""@ENGCHASISMMC"" 
                                        //                SET ""U_Status"" = 'Completed' 
                                        //                WHERE ""U_ProdOrdNo"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                        //                        rsUpdate.DoQuery(updateEngineChasis);

                                        //                        string updateProductionOrder = $@"UPDATE ""@ENGCHASPO"" SET ""U_Status"" = 'Completed', ""U_TotalTime"" = '{finalTotalTime}'  WHERE ""U_ProdDocEntry"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                        //                        rsUpdate.DoQuery(updateProductionOrder);


                                        //                        //SBO_Application.StatusBar.SetText("Status updated successfully in Engine Chasis Mapping Master and Production Order.",
                                        //                        //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        //                        //SBO_Application.MessageBox("Status updated successfully in Engine Chasis Mapping Master and Production Order.");
                                        //                    }
                                        //                    catch (Exception ex)
                                        //                    {
                                        //                        SBO_Application.StatusBar.SetText("Error updating Engine/Chasis: " + ex.Message,
                                        //                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //                        //SBO_Application.MessageBox("Error updating Engine/Chasis: " + ex.Message);
                                        //                    }
                                        //                }
                                        //                //SBO_Application.StatusBar.SetText("Job Stopped Successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        //                //SBO_Application.MessageBox("Job Stopped Successfully.");
                                        //                ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");

                                        //            }
                                        //            catch (Exception ex)
                                        //            {
                                        //                SBO_Application.StatusBar.SetText("Job Stopped, but failed to update details: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //                //SBO_Application.MessageBox("Job Stopped, but failed to update details: " + ex.Message);
                                        //            }
                                        //        }
                                        //        else
                                        //        {
                                        //            // Goods Issue Failed
                                        //            //SBO_Application.StatusBar.SetText("Failed to create Issue for Production. Job Status NOT changed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //            SBO_Application.MessageBox("Failed to create Issue for Production. Job Status NOT changed.");
                                        //            return;
                                        //        }
                                        //    }
                                        //    catch(Exception ex) { }
                                        //    finally
                                        //    {
                                        //        if (pb != null)
                                        //        {
                                        //            pb.Stop();
                                        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                        //            pb = null;
                                        //            scancount = 0;
                                        //            _scanBuffer = "";
                                        //            BubbleEvent = false;
                                        //        }
                                        //    }
                                        //}

                                        //else if (matchedColumn == "U_LastBarcode")
                                        //{
                                        //    SAPbouiCOM.ProgressBar pb = null;
                                        //    try
                                        //    {
                                        //        pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);
                                        //            try
                                        //            {
                                        //                // Create the Goods Issue 
                                        //                int goodIssue = Create_IssueFor_Production(
                                        //            _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan, _RSRoute.Fields.Item("U_JobExeNo").Value.ToString());

                                        //                if (goodIssue > 0)
                                        //                {
                                        //                    System.Threading.Thread.Sleep(200);
                                        //                    // Show the visual confirmation to the user.
                                        //                    ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp,
                                        //                        DateTime.Now.ToString("HH:mm:ss"), "#00e676",
                                        //                        "ACTION: Operation closed. Issue for Production Created.");
                                        //                }
                                        //                else
                                        //                {
                                        //                    SBO_Application.MessageBox("Failed to create Issue for Production. Status not updated.");
                                        //                }
                                        //            }
                                        //            catch (Exception ex)
                                        //            {
                                        //                // Log error
                                        //            }
                                        //    }
                                        //    catch (Exception ex)
                                        //    {
                                        //        _Application.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //    }
                                        //    finally
                                        //    {
                                        //        if (pb != null)
                                        //        {
                                        //            pb.Stop();
                                        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                        //            pb = null;
                                        //        }
                                        //        // These cleanup variables are important
                                        //        scancount = 0;
                                        //        _scanBuffer = "";
                                        //    }
                                        //}

                                        else if (matchedColumn == "U_LastBarcode")
                                        {
                                            if (_isProcessingIssue)
                                            {
                                                SBO_Application.StatusBar.SetText("System is busy processing a previous scan. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                return;
                                            }

                                            _isProcessingIssue = true;


                                            // 1. Capture UI variables before entering the background thread
                                            string docEntryToFind = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();
                                            string cNo = FirstScan;
                                            string mName = model;
                                            string pName = productName;
                                            string rDsp = routeDsp;
                                            string prdEnt = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString();
                                            string rId = routeId;
                                            string wId = WorkId;
                                            string jExe = _RSRoute.Fields.Item("U_JobExeNo").Value.ToString();

                                            // 2. Immediate UI cleanup
                                            scancount = 0;
                                            _scanBuffer = "";
                                            ShowProcessingMessage(cNo, mName, pName, rDsp);

                                            // 3. Process DI-API in background (to prevent SAP from hanging)
                                            //System.Threading.Tasks.Task.Run(() =>
                                            //{
                                            Task.Run(async () => 
                                            {
                                                try
                                                {
                                                    // This triggers your HANA Post-Transaction Stored Proc
                                                    //int goodIssue = Create_IssueFor_Production(prdEnt, rId, wId, cNo, jExe);
                                                    int goodIssue = await Create_Issue_ServiceLayer(prdEnt, rId, wId, cNo, jExe);

                                                    if (goodIssue > 0)
                                                    {
                                                        // IMPORTANT: Wait for HANA to finish the Update query (Post Transaction)
                                                        // 600ms is the "Sweet Spot" for HANA to commit the transaction
                                                        System.Threading.Thread.Sleep(600);

                                                        // 4. Open the form on the main thread
                                                        // Using LoadFinishedJobUI ensures it happens AFTER the DB update
                                                        LoadFinishedJobUI(docEntryToFind);

                                                        // Show browser success message
                                                        ShowBrowserMessage("✅", "JOB FINISHED", cNo, mName, pName, rDsp,
                                                            DateTime.Now.ToString("HH:mm:ss"), "#00e676",
                                                            "Operation closed. Post-Transaction sync complete.", true);
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Utilities.Application.SBO_Application.StatusBar.SetText("Error in background task: " + ex.Message);
                                                }
                                                finally
                                                {
                                                    _isProcessingIssue = false;
                                                }
                                            });

                                            BubbleEvent = false;
                                            return;
                                        }

                                        else
                                        {
                                            ShowBrowserMessage("❌", "INVALID ACTION", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is currently RUNNING. Only PAUSE or STOP allowed.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Start'. You can only PAUSE or STOP.");
                                            _scanBuffer = "";
                                            scancount = 0; // Reset
                                            BubbleEvent = false;
                                            return;
                                        }

                                    }
                                    #endregion

                                    #region When current status is "Pause" (User can scan Resume BarCode)
                                    else if (currentStatus == "Pause")
                                    {
                                        if (matchedColumn == "U_JobResume")
                                        {
                                            oGeneralData.SetProperty("U_Status", "Resume");

                                            // We only need to update the LAST row (the one currently open)
                                            if (oChildren.Count > 0)
                                            {
                                                // Get the last row added
                                                oChild = oChildren.Item(oChildren.Count - 1);

                                                DateTime manualTime = new DateTime(2025, 12, 08, 10, 40, 00);
                                                oChild.SetProperty("U_BrkEndDt", now);
                                                oChild.SetProperty("U_BrkEndTm", now);

                                                // Calculate duration for THIS specific break
                                                string sDt = Convert.ToString(oChild.GetProperty("U_BrkStartDt"));
                                                string sTm = Convert.ToString(oChild.GetProperty("U_BrkStartTm"));

                                                //DateTime startDT;
                                                //if (TryGetDateTime(sDt, sTm, out startDT))
                                                //{
                                                //    double minutes = (now - startDT).TotalMinutes;

                                                //    // Store line duration
                                                //    TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                //    DateTime lineDuration = new DateTime(1899, 12, 30).Add(ts);
                                                //    oChild.SetProperty("U_BrkTime", lineDuration);
                                                //}
                                                if (sDt != null && sTm != null)
                                                {
                                                    DateTime fullStart = GetFullDateTime(sDt, sTm);
                                                    DateTime fullEnd = now; // Since we just set it to manualTime

                                                    // This subtraction (fullEnd - fullStart) automatically handles 
                                                    // date changes (e.g. Dec 7 to Dec 8)
                                                    double minutes = (fullEnd - fullStart).TotalMinutes;

                                                    if (minutes < 0) minutes = 0; // Safety

                                                    //TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                    //DateTime lineDuration = new DateTime(1899, 12, 30).Add(ts);
                                                    //oChild.SetProperty("U_BrkTime", lineDuration);

                                                    TimeSpan ts = TimeSpan.FromMinutes(minutes);
                                                    string breakTimeStr = string.Format("{0:00}:{1:00}", (int)ts.TotalHours, ts.Minutes);
                                                    oChild.SetProperty("U_BrkTime", breakTimeStr);
                                                }

                                                 totalBreakMinutes = 0;

                                                for (int i = 0; i < oChildren.Count; i++)
                                                {
                                                    oChild = oChildren.Item(i);
                                                    object oSDt = oChild.GetProperty("U_BrkStartDt");
                                                    object oSTm = oChild.GetProperty("U_BrkStartTm");
                                                    object oEDt = oChild.GetProperty("U_BrkEndDt");
                                                    object oETm = oChild.GetProperty("U_BrkEndTm");

                                                    if (oSDt != null && oSTm != null && oEDt != null && oETm != null)
                                                    {
                                                        try
                                                        {
                                                            DateTime rowStart = GetFullDateTime(oSDt, oSTm);
                                                            DateTime rowEnd = GetFullDateTime(oEDt, oETm);
                                                            double rowMinutes = (rowEnd - rowStart).TotalMinutes;

                                                            if (rowMinutes > 0)
                                                            {
                                                                // Format individual row break time as HH:mm (supports > 24h)
                                                                int rHrs = (int)(rowMinutes / 60);
                                                                int rMns = (int)(rowMinutes % 60);
                                                                string rowDuration = $"{rHrs:D2}:{rMns:D2}";

                                                                // If U_BrkTime is also Alpha, set as string. If it's still Time, keep old logic.
                                                                oChild.SetProperty("U_BrkTime", rowDuration);

                                                                totalBreakMinutes += rowMinutes;
                                                            }
                                                        }
                                                        catch (Exception ex) { }
                                                    }
                                                }

                                                // Set the Grand Total Break Time in the Header as a String
                                                int bHrs = (int)(totalBreakMinutes / 60);
                                                int bMns = (int)(totalBreakMinutes % 60);
                                                totalBrkDurationString = $"{bHrs:D2}:{bMns:D2}";
                                                oGeneralData.SetProperty("U_TotBrkTime", totalBrkDurationString);


                                                //oForm = Utilities.Application.SBO_Application.Forms.Item("frmWJobEx");
                                                //_Matrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mtBreak").Specific;
                                                //((SAPbouiCOM.EditTextColumn)_Matrix.Columns.Item("colBrkTm")).ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Manual;
                                            }

                                            //SBO_Application.StatusBar.SetText("Job Resumed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Resumed");
                                            //ShowBrowserMessage("▶️", "JOB RESUMED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly restarted. Continue work.");
                                            ShowBrowserMessage("✅", "JOB RESUMED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly restarted. Continue work.");
                                        }
                                        else
                                        {
                                            ShowBrowserMessage("❌", "RESUME REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is PAUSED. You must scan RESUME to proceed.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Pause'. You must scan RESUME to continue.");
                                            scancount = 0;
                                            _scanBuffer = "";
                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                    #endregion

                                    #region When Current Status is "Resumed" then two scenario's (Pause and Stop))
                                    else if (currentStatus == "Resume")
                                    {
                                        #region  Pause scanned(add a new line)   
                                        if (matchedColumn == "U_JobPause")
                                        {
                                            // Change Status to Pause
                                            oGeneralData.SetProperty("U_Status", "Pause");

                                            // Safety Check: Ensure the PREVIOUS break row is actually closed.
                                            if (oChildren.Count > 0)
                                            {
                                                SAPbobsCOM.GeneralData oLastChild = oChildren.Item(oChildren.Count - 1);
                                                string lastEndDt = Convert.ToString(oLastChild.GetProperty("U_BrkEndDt"));

                                                // If the last row has no End Date, close it now before starting a new one
                                                if (string.IsNullOrEmpty(lastEndDt))
                                                {
                                                    
                                                    oLastChild.SetProperty("U_BrkEndDt", now);
                                                    oLastChild.SetProperty("U_BrkEndTm", now);
                                                    // Calculate time for this "forgotten" break here if needed
                                                }
                                            }

                                            // ADD A NEW ROW for the new Pause
                                            oChild = oChildren.Add(); 
                                            oChild.SetProperty("U_BrkStartDt", now);
                                            oChild.SetProperty("U_BrkStartTm", now);

                                            //SBO_Application.StatusBar.SetText("Job Paused again. New break line added.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            //SBO_Application.MessageBox("Job Paused again. New break line added.");
                                            //ShowBrowserMessage("✅ Success", "Job Paused Sucessfully", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#28a745");
                                            ShowBrowserMessage("✅", "JOB PAUSED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Job is PAUSED. You must scan RESUME to proceed.");
                                            oGeneralService.Update(oGeneralData);
                                            this._Object = new ClsJobOrderExecution();
                                            Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);

                                            // 2. Get a reference to the SAP Form from your newly created object
                                            // Note: This assumes your ClsJobOrderExecution class has a public 'Form' or 'oForm' property
                                            SAPbouiCOM.Form oForm = ((ClsJobOrderExecution)this._Object).Form;

                                            if (oForm != null)
                                            {
                                                oForm.Freeze(true);
                                                try
                                                {
                                                    // 3. Switch the form to Find Mode
                                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                                    // 4. Set the DocEntry into your field (using the UID you mentioned: "txtDocE")
                                                    // If "txtDocE" is the UID on your custom .srf, this will work.
                                                    ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocE").Specific).Value = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();

                                                    // 5. Simulate clicking the "Find" button (UID "1")
                                                    oForm.Items.Item("1").Click();

                                                    SetRemarkFieldsState(oForm, "Pause");
                                                }
                                                catch (Exception ex)
                                                {
                                                    Utilities.Application.SBO_Application.StatusBar.SetText("Failed to find record: " + ex.Message);
                                                }
                                                finally
                                                {
                                                    oForm.Freeze(false);
                                                }
                                            }
                                        }
                                        #endregion

                                        #region Final Stop BarCode Scanned After Resume Status (Calculation of all Breaks and creation of goods issue)
                                        //else if (matchedColumn == "U_LastBarcode")
                                        //{
                                        //    SAPbouiCOM.ProgressBar pb = null;
                                        //    try
                                        //    {
                                        //        pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);

                                        //        int goodIssue = Create_IssueFor_Production(_RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan, _RSRoute.Fields.Item("U_JobExeNo").Value.ToString());
                                        //        //int goodIssue = 786;

                                        //        // Only proceed if Goods Issue was successful
                                        //        if (goodIssue > 0)
                                        //        {
                                        //            oGeneralData.SetProperty("U_Status", "Stop");
                                        //            oGeneralData.SetProperty("U_EndDate", now);
                                        //            oGeneralData.SetProperty("U_EndTime", now);
                                        //            oGeneralData.SetProperty("U_IForProdEntry", goodIssue.ToString());


                                        //            //double totalBreakMinutes = 0;

                                        //            //for (int i = 0; i < oChildren.Count; i++)
                                        //            //{
                                        //            //    oChild = oChildren.Item(i);
                                        //            //    object oSDt = oChild.GetProperty("U_BrkStartDt");
                                        //            //    object oSTm = oChild.GetProperty("U_BrkStartTm");
                                        //            //    object oEDt = oChild.GetProperty("U_BrkEndDt");
                                        //            //    object oETm = oChild.GetProperty("U_BrkEndTm");

                                        //            //    if (oSDt != null && oSTm != null && oEDt != null && oETm != null)
                                        //            //    {
                                        //            //        try
                                        //            //        {
                                        //            //            DateTime rowStart = GetFullDateTime(oSDt, oSTm);
                                        //            //            DateTime rowEnd = GetFullDateTime(oEDt, oETm);
                                        //            //            double rowMinutes = (rowEnd - rowStart).TotalMinutes;

                                        //            //            if (rowMinutes > 0)
                                        //            //            {
                                        //            //                // Format individual row break time as HH:mm (supports > 24h)
                                        //            //                int rHrs = (int)(rowMinutes / 60);
                                        //            //                int rMns = (int)(rowMinutes % 60);
                                        //            //                string rowDuration = $"{rHrs:D2}:{rMns:D2}";

                                        //            //                // If U_BrkTime is also Alpha, set as string. If it's still Time, keep old logic.
                                        //            //                oChild.SetProperty("U_BrkTime", rowDuration);

                                        //            //                totalBreakMinutes += rowMinutes;
                                        //            //            }
                                        //            //        }
                                        //            //        catch (Exception ex) { }
                                        //            //    }
                                        //            //}

                                        //            //// Set the Grand Total Break Time in the Header as a String
                                        //            //int bHrs = (int)(totalBreakMinutes / 60);
                                        //            //int bMns = (int)(totalBreakMinutes % 60);
                                        //            //string totalBrkDurationString = $"{bHrs:D2}:{bMns:D2}";
                                        //            //oGeneralData.SetProperty("U_TotBrkTime", totalBrkDurationString);

                                        //            try
                                        //            {

                                        //                DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                        //                DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                        //                double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                        //                double totalActMinutes = totalElapsedMinutes - totalBreakMinutes;

                                        //                int hours = (int)(totalActMinutes / 60);
                                        //                int mins = (int)(totalActMinutes % 60);
                                        //                string durationString = $"{hours:D2}:{mins:D2}";


                                        //                //if alpha
                                        //                //oGeneralData.SetProperty("U_TotActTime", durationString);
                                        //                oGeneralData.SetProperty("U_TotActTime", durationString);

                                        //                // NOTE: Visually in SAP, this will show "01:16", but the underlying value 
                                        //                // will contain the extra day (1899-12-31) which is important for the next step.
                                        //                //oGeneralData.SetProperty("U_TotActTime", new DateTime(1899, 12, 30).AddMinutes(totalActMinutes));
                                        //            }
                                        //            catch (Exception ex)
                                        //            {
                                        //            }


                                        //            // Update External Tables
                                        //            try
                                        //            {

                                        //                DateTime dtStartDate = (DateTime)oGeneralData.GetProperty("U_StartDate");
                                        //                DateTime dtStartTime = (DateTime)oGeneralData.GetProperty("U_StartTime");
                                        //                DateTime dtEndDate = (DateTime)oGeneralData.GetProperty("U_EndDate");
                                        //                DateTime dtEndTime = (DateTime)oGeneralData.GetProperty("U_EndTime");
                                        //                //DateTime dtBrkTime = (DateTime)oGeneralData.GetProperty("U_TotBrkTime");
                                        //                //DateTime dtActTime = (DateTime)oGeneralData.GetProperty("U_TotActTime");
                                        //                string strBrkTime = Convert.ToString(oGeneralData.GetProperty("U_TotBrkTime"));
                                        //                string strActTime = Convert.ToString(oGeneralData.GetProperty("U_TotActTime"));


                                        //                string sqlStartDate = dtStartDate.ToString("yyyyMMdd");
                                        //                string sqlEndDate = dtEndDate.ToString("yyyyMMdd");
                                        //                string sqlStartTime = dtStartTime.ToString("HHmm");
                                        //                string sqlEndTime = dtEndTime.ToString("HHmm");
                                        //                //string sqlBrkTime = dtBrkTime.ToString("HHmm");
                                        //                //string sqlActTime = dtActTime.ToString("HHmm");
                                        //                string sqlBrkTime = strBrkTime.Replace(":", "").PadLeft(4, '0');
                                        //                string sqlActTime = strActTime.Replace(":", "").PadLeft(4, '0');

                                        //                /// for exceeding 24 hours 
                                        //                DateTime dtHeaderStart = GetFullDateTime(oGeneralData.GetProperty("U_StartDate"), oGeneralData.GetProperty("U_StartTime"));
                                        //                DateTime dtLastChildEnd = GetFullDateTime(oGeneralData.GetProperty("U_EndDate"), oGeneralData.GetProperty("U_EndTime"));

                                        //                double totalElapsedMinutes = (dtLastChildEnd - dtHeaderStart).TotalMinutes;
                                        //                double totalActMinutes = totalElapsedMinutes - totalBreakMinutes;

                                        //                int totalHrs = (int)(totalActMinutes / 60);
                                        //                int totalMns = (int)(totalActMinutes % 60);

                                        //                // Use a string format for the SQL update
                                        //                // This will result in "2516" instead of "0116"
                                        //                string manualTimeStr = $"{totalHrs:D2}:{totalMns:D2}";
                                        //                /// end for exceeding 24 hours


                                        //                SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                        //                //string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                        //                //    SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                        //                //    ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                        //                //    ""U_TotalBrkDwnT"" = '{sqlBrkTime}', ""U_ActualTimeCon"" = '{sqlActTime}'
                                        //                //    WHERE ""U_WorkId"" = '{WorkId}'";
                                        //                //rsUpdate.DoQuery(updateQuery);

                                        //                //for 24 hours exceed
                                        //                string updateQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                        //                SET ""U_IssueForProdEntry"" = '{goodIssue}' , ""U_Status"" = 'Closed', ""U_StartDate"" = '{sqlStartDate}',
                                        //                ""U_StartTime"" = '{sqlStartTime}', ""U_EndDate"" = '{sqlEndDate}', ""U_EndTime"" = '{sqlEndTime}', 
                                        //                ""U_TotalBrkDwnT"" = '{totalBrkDurationString}', ""U_ActualTimeCon"" = '{manualTimeStr}'
                                        //                WHERE ""U_WorkId"" = '{WorkId}'";
                                        //                rsUpdate.DoQuery(updateQuery);
                                        //                /////////

                                        //                string chasisNo = oGeneralData.GetProperty("U_ChassisNo");
                                        //                string selectQuery = $@"SELECT * FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}' and ""U_IssueForProdEntry"" IS NULL ";
                                        //                SAPbobsCOM.Recordset rsSelect = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        //                rsSelect.DoQuery(selectQuery);

                                        //                if (rsSelect.RecordCount > 0)
                                        //                {
                                        //                    SBO_Application.StatusBar.SetText("All processes must be completed before updating Engine/Chasis.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //                }
                                        //                else
                                        //                {
                                        //                    try
                                        //                    {
                                        //                        int totalMinutes = 0;
                                        //                        string sumQuery = $@"SELECT ""U_ActualTimeCon"" FROM ""@WRKORDRDTLSC"" WHERE ""U_ChasisNo"" = '{chasisNo}'";

                                        //                        rsSelect.DoQuery(sumQuery);

                                        //                        while (!rsSelect.EoF)
                                        //                        {
                                        //                            string timeStr = rsSelect.Fields.Item("U_ActualTimeCon").Value.ToString().Replace(":", "").Trim();

                                        //                            if (!string.IsNullOrEmpty(timeStr))
                                        //                            {
                                        //                                // Pad to ensure we have at least HHmm (e.g., "15" -> "0015", "130" -> "0130")
                                        //                                timeStr = timeStr.PadLeft(4, '0');

                                        //                                // Take the last 4 digits only to be safe
                                        //                                if (timeStr.Length > 4)
                                        //                                    timeStr = timeStr.Substring(timeStr.Length - 4);

                                        //                                if (int.TryParse(timeStr.Substring(0, 2), out int hours) &&
                                        //                                    int.TryParse(timeStr.Substring(2, 2), out int mins))
                                        //                                {
                                        //                                    totalMinutes += (hours * 60) + mins;
                                        //                                }
                                        //                            }
                                        //                            rsSelect.MoveNext();
                                        //                        }

                                        //                        // Convert total minutes back to HHmm format
                                        //                        int finalHours = totalMinutes / 60;
                                        //                        int finalMins = totalMinutes % 60;
                                        //                        string finalTotalTime = $"{finalHours:00}:{finalMins:00}";

                                        //                        //Update 
                                        //                        string updateEngineChasis = $@"UPDATE ""@ENGCHASISMMC"" 
                                        //                SET ""U_Status"" = 'Completed' 
                                        //                WHERE ""U_ProdOrdNo"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                        //                        rsUpdate.DoQuery(updateEngineChasis);

                                        //                        string updateProductionOrder = $@"UPDATE ""@ENGCHASPO"" SET ""U_Status"" = 'Completed', ""U_TotalTime"" = '{finalTotalTime}'  WHERE ""U_ProdDocEntry"" = '{productionOrderEntry}' and ""U_ChasisNo"" = '{chasisNo}'";
                                        //                        rsUpdate.DoQuery(updateProductionOrder);


                                        //                        //SBO_Application.StatusBar.SetText("Status updated successfully in Engine Chasis Mapping Master and Production Order.",
                                        //                        //    SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                        //                        //SBO_Application.MessageBox("Status updated successfully in Engine Chasis Mapping Master and Production Order.");
                                        //                    }
                                        //                    catch (Exception ex)
                                        //                    {
                                        //                        SBO_Application.StatusBar.SetText("Error updating Engine/Chasis: " + ex.Message,
                                        //                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //                        //SBO_Application.MessageBox("Error updating Engine/Chasis: " + ex.Message);
                                        //                    }
                                        //                }
                                        //            }
                                        //            catch (Exception ex)
                                        //            {
                                        //                SBO_Application.StatusBar.SetText("Job Stopped, but failed to update details: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //            }
                                        //            ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");
                                        //        }
                                        //        else
                                        //        {
                                        //            //SBO_Application.StatusBar.SetText("Failed to create Issue for Production. Job NOT stopped.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //            SBO_Application.MessageBox("Failed to create Issue for Production. Job NOT stopped.");
                                        //            return; // Exit
                                        //        }
                                        //    }
                                        //    catch(Exception ex)
                                        //    {
                                        //        SBO_Application.MessageBox("Critical Error: " + ex.Message);
                                        //    }
                                        //    finally
                                        //    {
                                        //        if (pb != null)
                                        //        {
                                        //            pb.Stop();
                                        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                        //            pb = null;
                                        //            scancount = 0;
                                        //            _scanBuffer = "";
                                        //            BubbleEvent = false;
                                        //        }
                                        //        //pb.Stop();
                                        //        //ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");
                                        //    }

                                        //}

                                        //else if (matchedColumn == "U_LastBarcode")
                                        //{
                                        //    SAPbouiCOM.ProgressBar pb = null;
                                        //    try
                                        //    {
                                        //        pb = SBO_Application.StatusBar.CreateProgressBar("Creating Production Issue........", 1, false);

                                        //        // Create the Goods Issue only.
                                        //        int goodIssue = Create_IssueFor_Production(_RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString(), routeId, WorkId, FirstScan, _RSRoute.Fields.Item("U_JobExeNo").Value.ToString());

                                        //        if (goodIssue > 0)
                                        //        {

                                        //            // Show the visual confirmation to the user.
                                        //            System.Threading.Thread.Sleep(200);
                                        //            ShowBrowserMessage("✅", "JOB FINISHED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Operation closed. Issue for Production Created.");
                                        //        }
                                        //        else
                                        //        {
                                        //            SBO_Application.MessageBox("Failed to create Issue for Production. Status not updated.");
                                        //        }
                                        //    }
                                        //    catch (Exception ex)
                                        //    {
                                        //        _Application.StatusBar.SetText("Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //    }
                                        //    finally
                                        //    {
                                        //        if (pb != null)
                                        //        {
                                        //            pb.Stop();
                                        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                                        //            pb = null;
                                        //        }
                                        //        // These cleanup variables are important
                                        //        scancount = 0;
                                        //        _scanBuffer = "";
                                        //    }
                                        //}

                                        else if (matchedColumn == "U_LastBarcode")
                                        {
                                            if (_isProcessingIssue)
                                            {
                                                SBO_Application.StatusBar.SetText("System is busy processing a previous scan. Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                return;
                                            }

                                            _isProcessingIssue = true;

                                            // 1. Capture UI variables before entering the background thread
                                            string docEntryToFind = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();
                                            string cNo = FirstScan;
                                            string mName = model;
                                            string pName = productName;
                                            string rDsp = routeDsp;
                                            string prdEnt = _RSRoute.Fields.Item("U_PrdOrdEnt").Value.ToString();
                                            string rId = routeId;
                                            string wId = WorkId;
                                            string jExe = _RSRoute.Fields.Item("U_JobExeNo").Value.ToString();

                                            // 2. Immediate UI cleanup
                                            scancount = 0;
                                            _scanBuffer = "";
                                            ShowProcessingMessage(cNo, mName, pName, rDsp);

                                            // 3. Process DI-API in background (to prevent SAP from hanging)
                                            //System.Threading.Tasks.Task.Run(() =>
                                            //{
                                            Task.Run(async () =>
                                            {
                                                try
                                                {
                                                    // This triggers your HANA Post-Transaction Stored Proc
                                                    //int goodIssue = Create_IssueFor_Production(prdEnt, rId, wId, cNo, jExe);
                                                    int goodIssue = await Create_Issue_ServiceLayer(prdEnt, rId, wId, cNo, jExe);

                                                    if (goodIssue > 0)
                                                    {
                                                        // IMPORTANT: Wait for HANA to finish the Update query (Post Transaction)
                                                        System.Threading.Thread.Sleep(600);

                                                        // Open the form on the main thread
                                                        // Using LoadFinishedJobUI ensures it happens AFTER the DB update
                                                        LoadFinishedJobUI(docEntryToFind);

                                                        // Show browser success message
                                                        ShowBrowserMessage("✅", "JOB FINISHED", cNo, mName, pName, rDsp,
                                                            DateTime.Now.ToString("HH:mm:ss"), "#00e676",
                                                            "Operation closed. Post-Transaction sync complete.", true);
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Utilities.Application.SBO_Application.StatusBar.SetText("Error in background task: " + ex.Message);
                                                }
                                                finally
                                                {
                                                    _isProcessingIssue = false;
                                                }
                                            });

                                            BubbleEvent = false;
                                            return;
                                        }
                                        #endregion

                                        else
                                        {
                                            ShowBrowserMessage("❌", "PAUSE/ STOP REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "INVALID ACTION: You can only PAUSE or STOP.");
                                            //SBO_Application.MessageBox($"Invalid Action! Current status is 'Resume'. You can only PAUSE or STOP.");
                                            scancount = 0;
                                            _scanBuffer = "";
                                            BubbleEvent = false;
                                            return; // Scanned something else while in Resume
                                        }
                                    }
                                    #endregion

                                    #region When Current Status is "Stop" dont let to do anything
                                    else if (currentStatus == "Stop")
                                    {
                                        //SBO_Application.StatusBar.SetText("Job already stopped!", SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        //SBO_Application.MessageBox("Job already stopped!");
                                        ShowBrowserMessage("🔔", "ALREADY COMPLETED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#17a2b8", "ACTION: This stage is closed. No further actions allowed.");
                                    }
                                    #endregion

                                    #region Try to reopen the Job Execution Form
                                    //// --- Final Commit ---
                                    oGeneralService.Update(oGeneralData);

                                    //string docEntry = _RSWorkExecDetails.Fields.Item("DocEntry").Value.ToString();
                                    //string uniqueFormID = "WJOBEXE_" + docEntry;

                                    //try
                                    //{
                                    //    bool formFound = false;

                                    //    // Check if this specific Job Form is ALREADY open
                                    //    try
                                    //    {
                                    //        // Try to grab the form by its Unique ID
                                    //        SAPbouiCOM.Form existingForm = SBO_Application.Forms.Item(uniqueFormID);
                                    //        existingForm.Select(); // Bring it to the front

                                    //        // Refresh to show the new status 
                                    //        existingForm.Refresh();

                                    //        formFound = true;
                                    //    }
                                    //    catch
                                    //    {
                                    //        formFound = false;
                                    //    }

                                    //    // If the form is NOT open, we create it and load the data
                                    //    if (!formFound)
                                    //    {
                                    //        SAPbouiCOM.FormCreationParams cp = (SAPbouiCOM.FormCreationParams)SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                                    //        cp.FormType = "WJOBEXE";  
                                    //        cp.ObjectType = "WJOBEXE"; 
                                    //        cp.UniqueID = uniqueFormID;

                                    //        SAPbouiCOM.Form oForm = SBO_Application.Forms.AddEx(cp);

                                    //        // Safely switch to Find Mode and Load the Record
                                    //        oForm.Freeze(true);
                                    //        try
                                    //        {
                                    //            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                    //            // IMPORTANT: "0_U_E" is usually the default UI ID for the DocEntry/DocNum field in a UDO.
                                    //            // We use the DocEntry from the DB (docEntry variable) to find the record exact match.
                                    //            SAPbouiCOM.EditText txtFind = (SAPbouiCOM.EditText)oForm.Items.Item("0_U_E").Specific;
                                    //            txtFind.Value = docEntry;

                                    //            // Click the Find Button (Item "1")
                                    //            oForm.Items.Item("1").Click();
                                    //        }
                                    //        finally
                                    //        {
                                    //            oForm.Freeze(false);
                                    //        }
                                    //    }
                                    //}
                                    //catch (Exception ex)
                                    //{
                                    //    SBO_Application.StatusBar.SetText(
                                    //        "Job updated, but error opening form: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                                    //    );
                                    //}
                                    #endregion
                                }
                                else
                                {
                                    string newDocEntry = "";
                                    if (matchedColumn == "U_FirstBarcode")
                                    {

                                        // 1. Initialize Company Service and General Service
                                        // Note: oCompany must be your valid SAP Company object
                                        SAPbobsCOM.CompanyService oCompanyService = Utilities.Application.Company.GetCompanyService();
                                        SAPbobsCOM.GeneralService oGeneralService;
                                        SAPbobsCOM.GeneralData oGeneralData;
                                        SAPbobsCOM.GeneralDataParams oGeneralParams;

                                        try
                                        {
                                            string query_WorkOrderD = $@"SELECT T0.""U_ProdOrdNo"",T0.""U_ProductCode"",T0.""U_AdvancedSBNo"",T0.""U_ProductName"",T0.""U_AdvancedSBNo"",T0.""U_DocNo"",
                                                             T1.""U_EngineNo"",T1.""U_ChasisNo"",T1.""U_BatchNo"",T1.""U_WorkId"",T1.""U_RouteDsp"", T1.""U_JobExeNo""
                                                            from ""@WRKORDRDTLSH""  T0 
                                                            inner join ""@WRKORDRDTLSC""  T1 on T0.""DocEntry"" = T1.""DocEntry""  
                                                            WHERE T1.""U_WorkId"" = '{WorkId}'";

                                            SAPbobsCOM.Recordset _RSWorkOrderDetails = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                            _RSWorkOrderDetails.DoQuery(query_WorkOrderD);
                                            oGeneralService = oCompanyService.GetGeneralService("WJOBEXE");
                                            oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                                            // --- Populate UDO fields directly from Recordset values ---
                                            oGeneralData.SetProperty("U_JobID", WorkId);
                                            oGeneralData.SetProperty("U_DocNo", Utilities.getMaxColumnValueNum("@WJOBEXEH", "U_DocNo"));
                                            oGeneralData.SetProperty("U_JobDesc", _RSWorkOrderDetails.Fields.Item("U_RouteDsp").Value.ToString());
                                            oGeneralData.SetProperty("U_ProdOrdNo", _RSWorkOrderDetails.Fields.Item("U_ProdOrdNo").Value.ToString());
                                            oGeneralData.SetProperty("U_WorkOrdNo", _RSWorkOrderDetails.Fields.Item("U_DocNo").Value.ToString());
                                            oGeneralData.SetProperty("U_FGCode", _RSWorkOrderDetails.Fields.Item("U_ProductCode").Value.ToString());
                                            oGeneralData.SetProperty("U_FGDesc", _RSWorkOrderDetails.Fields.Item("U_ProductName").Value.ToString());
                                            oGeneralData.SetProperty("U_EngineNo", _RSWorkOrderDetails.Fields.Item("U_EngineNo").Value.ToString());
                                            oGeneralData.SetProperty("U_ChassisNo", _RSWorkOrderDetails.Fields.Item("U_ChasisNo").Value.ToString());
                                            oGeneralData.SetProperty("U_BatchNo", _RSWorkOrderDetails.Fields.Item("U_BatchNo").Value.ToString());
                                            oGeneralData.SetProperty("U_Status", "Start");

                                            // Auto Start Date & Time
                                            oGeneralData.SetProperty("U_StartDate", DateTime.Now);
                                            oGeneralData.SetProperty("U_StartTime", DateTime.Now);
                                            oGeneralParams = oGeneralService.Add(oGeneralData);
                                            newDocEntry = oGeneralParams.GetProperty("DocEntry").ToString();
                                            // Operator from RS (or current user)
                                            //oGeneralData.SetProperty("U_Operator", _RSWorkExecDetails.Fields.Item("U_Operator").Value.ToString());

                                            // Update Work Order Details Document with recently created DocEntry
                                            string prodOrdNo = _RSWorkOrderDetails.Fields.Item("U_ProdOrdNo").Value.ToString();
                                            string chasisNo = _RSWorkOrderDetails.Fields.Item("U_ChasisNo").Value.ToString();
                                            if (!string.IsNullOrEmpty(newDocEntry))
                                            {
                                                try
                                                {
                                                    SAPbobsCOM.Recordset rsUpdateRef = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string updateRefQuery = $@"UPDATE ""@WRKORDRDTLSC"" 
                                                       SET ""U_JobExeNo"" = {newDocEntry} , ""U_Status"" = 'WIP'
                                                       WHERE ""U_WorkId"" = '{WorkId}'";

                                                    rsUpdateRef.DoQuery(updateRefQuery);

                                                    SAPbobsCOM.Recordset rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                                    string updateQuery = $@"UPDATE ""@ENGCHASPO"" 
                                                       SET  ""U_Status"" = 'Started'
                                                       WHERE ""U_ProdDocEntry"" = '{prodOrdNo}' and ""U_ChasisNo"" = '{chasisNo}' ";

                                                    rsUpdate.DoQuery(updateQuery);
                                                }
                                                catch (Exception exUpdate)
                                                {
                                                    SBO_Application.StatusBar.SetText("Job created, but failed to link DocEntry in Work Order Details: " + exUpdate.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                                }
                                            }

                                        }
                                        catch (Exception ex)
                                        {
                                            // Handle Errors
                                            string errorMsg = ex.Message;
                                            // System.Windows.Forms.MessageBox.Show("Error creating Job: " + errorMsg);
                                        }
                                        finally
                                        {
                                            // Clean up COM objects if necessary (though GC handles them usually)
                                            //if (oGeneralData != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                                            //if (oGeneralService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                                        }
                                        //ShowBrowserMessage("🚀", "JOB STARTED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly initiated. Work in progress.");
                                        ShowBrowserMessage("✅", "JOB STARTED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#00e676", "ACTION: Assembly initiated. Work in progress.");

                                        //Utilities.Application.SBO_Application.ActivateMenuItem("WJOBEXE");
                                        //SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;
                                        //oForm.Freeze(true);
                                        //oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                        //((SAPbouiCOM.EditText)oForm.Items.Item("txtDocE").Specific).Value = newDocEntry;
                                        //oForm.Items.Item("1").Click();
                                        //oForm.Freeze(false);

                                        this._Object = new ClsJobOrderExecution();
                                        Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);

                                        // 2. Get a reference to the SAP Form from your newly created object
                                        // Note: This assumes your ClsJobOrderExecution class has a public 'Form' or 'oForm' property
                                        SAPbouiCOM.Form oForm = ((ClsJobOrderExecution)this._Object).Form;

                                        if (oForm != null)
                                        {
                                            oForm.Freeze(true);
                                            try
                                            {
                                                // 3. Switch the form to Find Mode
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                                // 4. Set the DocEntry into your field (using the UID you mentioned: "txtDocE")
                                                // If "txtDocE" is the UID on your custom .srf, this will work.
                                                ((SAPbouiCOM.EditText)oForm.Items.Item("txtDocE").Specific).Value = newDocEntry;

                                                // 5. Simulate clicking the "Find" button (UID "1")
                                                oForm.Items.Item("1").Click();

                                                SetRemarkFieldsState(oForm, "Start");
                                            }
                                            catch (Exception ex)
                                            {
                                                Utilities.Application.SBO_Application.StatusBar.SetText("Failed to find record: " + ex.Message);
                                            }
                                            finally
                                            {
                                                oForm.Freeze(false);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        string msg = "";
                                        if (string.IsNullOrEmpty(matchedColumn))
                                            msg = "Scanned barcode is not valid for this Chassis.";
                                        else
                                            msg = $"Cannot perform '{matchedColumn}' action. The Job has not been Started yet. Please scan the Start Barcode.";

                                        //SBO_Application.StatusBar.SetText(msg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        //SBO_Application.MessageBox(msg);
                                        ShowBrowserMessage("❌", "START REQUIRED", FirstScan, model, productName, routeDsp, DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Scan the START barcode to initiate this operation.");

                                        _scanBuffer = "";
                                        scancount = 0; 
                                        BubbleEvent = false;
                                    }
                                }
                            }
                            else
                            {
                                //SBO_Application.MessageBox("First Scan (Chassis) not found in Database.");
                                ShowBrowserMessage("❌", "DATA NOT FOUND", FirstScan, "", "", "", DateTime.Now.ToString("HH:mm:ss"), "#ff5252", "ACTION: Chassis mapping missing in Production Planning.");
                                scancount = 0;
                                _scanBuffer = "";
                                BubbleEvent = false;
                            }
                            scancount = 0;
                            _scanBuffer = "";
                            //addBtn.Enabled = true;
                            BubbleEvent = false;
                        }
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
 
        }

        #endregion

        #region RIGHT CLICK EVENT

        private void SboApplication_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (this._Collection.Contains(eventInfo.FormUID))
            {
                this._Object = (Production_Planning_LMC.Base)this._Collection[eventInfo.FormUID];
                this._Object.RightClick_Event(ref eventInfo, ref BubbleEvent);
            }
        }

        #endregion

        #region APPLICATION EVENT
        private void SboApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Utilities.LoadMenus(Constants.Menus.REMOVE_MENUS);
                        Utilities.ShowWarningMessage("Add-On is Disconnected Successfully ... ");
                        CloseApplication();
                        break;
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        #endregion

        #region FORM DATA EVENT
        private void SboApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.FormUID != null)
                {
                    if (this._Collection.Contains(BusinessObjectInfo.FormUID))
                    {
                        this._Object = (Production_Planning_LMC.Base)this._Collection[BusinessObjectInfo.FormUID];
                        this._Object.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        #endregion

        #region CLOSE APPLICATION
        private void CloseApplication()
        {
            try
            {
                if (this._Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Application);
                    _Application = null;
                }
                if (this._Company != null)
                {
                    if (this._Company.Connected)
                        this._Company.Disconnect();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company);
                    _Company = null;
                }
                this._Collection = null;
                this._LookCollection = null;
                _Listener = null;
                System.Windows.Forms.Application.Exit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // clsUtilities.oApplication = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        #endregion

        #region full Date and Time for calculating total break down time
        private DateTime GetFullDateTime(object oDate, object oTime)
        {
            // 1. Convert Objects to DateTime
            // SAP "Date" fields usually return "MM/dd/yyyy 00:00:00"
            // SAP "Time" fields usually return "1899/12/30 HH:mm:ss"

            DateTime dtDate = Convert.ToDateTime(oDate);
            DateTime dtTime = Convert.ToDateTime(oTime);

            // 2. Combine them: Take Year/Month/Day from Date, and Hour/Min/Sec from Time
            DateTime fullDateTime = new DateTime(
                dtDate.Year,
                dtDate.Month,
                dtDate.Day,
                dtTime.Hour,
                dtTime.Minute,
                dtTime.Second
            );

            return fullDateTime;
        }
        #endregion

        #region Create Issue For Production 
        private int Create_IssueFor_Production(string ProdEntry, string route, string workId, string chassisNo, string JobExeNo)
        {
            int newDocEntry = 0;
            //SAPbouiCOM.ProgressBar oProgBar = null;
            // Create Issue for Production (Inventory Gen Exit)
            SAPbobsCOM.Recordset rsBOM = (SAPbobsCOM.Recordset)
                   Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rsMap = (SAPbobsCOM.Recordset)
     Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            SAPbobsCOM.Documents oIssue = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
            // Check Serial Type
            SAPbobsCOM.Recordset rsSer = (SAPbobsCOM.Recordset)
                Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                if (string.IsNullOrWhiteSpace(ProdEntry))
                    throw new Exception("Production Order missing.");

                //oProgBar = SBO_Application.StatusBar.CreateProgressBar("Initializing Auto Issue...", 2, false);

                int prodOrderDocEntry = Convert.ToInt32(ProdEntry);

                // Fetch BOM (WOR1) for the given Production Order + Stage

                string bomQuery = $@"
                    SELECT  
                        T0.""ItemCode"",
                        T0.""PlannedQty"",
                        T0.""IssueType"",
                        T0.""wareHouse"",
                        T0.""BaseQty"" AS ""UnitReq"", T2.""U_SerialType"",
                        T0.""LineNum"", T2.""ManSerNum"", T1.""PlannedQty"" ""headerQty""
                    FROM WOR1 T0 
                    INNER JOIN OWOR T1 ON T0.""DocEntry"" = T1.""DocEntry""
                    INNER JOIN OITM T2 on T0.""ItemCode"" = T2.""ItemCode""
                    WHERE T0.""DocEntry"" = '{prodOrderDocEntry}'
                     AND T0.""StageId"" = '{route}' and T2.""InvntItem"" = 'Y' ";////'{prodOrderDocEntry}' AND T0.""StageId"" = '1' {route} ";

                rsBOM.DoQuery(bomQuery);

                //  Fetch Chassis Mapping from UDT

                string mapQuery = $@"
                SELECT * 
                FROM ""@ENGCHASISMMC""
                WHERE ""U_ChasisNo"" = '{chassisNo}'";

                rsMap.DoQuery(mapQuery);

                if (rsMap.EoF)
                    throw new Exception("Mapping not found for Chassis No " + chassisNo);


                oIssue.DocDate = DateTime.Now;
                oIssue.Comments = $"Auto Issue — STOP Trigger | Work: {workId}, Route: {route}";
                oIssue.UserFields.Fields.Item("U_ProdEntry").Value = prodOrderDocEntry;
                oIssue.UserFields.Fields.Item("U_JobExeNo").Value = JobExeNo;

                //oProgBar.Text = "Preparing Lines... Please wait.";

                //  Loop through BOM items
                while (!rsBOM.EoF)
                {
                    string itemCode = rsBOM.Fields.Item("ItemCode").Value.ToString();
                    double plannedQty = Convert.ToDouble(rsBOM.Fields.Item("PlannedQty").Value);
                    double unitReq = Convert.ToDouble(rsBOM.Fields.Item("UnitReq").Value);
                    string issueMethod = rsBOM.Fields.Item("IssueType").Value.ToString();
                    string whs = rsBOM.Fields.Item("wareHouse").Value.ToString();
                    int lineNum = Convert.ToInt32(rsBOM.Fields.Item("LineNum").Value);
                    //string isSerial = rsBOM.Fields.Item("ManSerNum").Value.ToString();
                    double headerPlannedQty = Convert.ToDouble(rsBOM.Fields.Item("headerQty").Value);

                    // Link Line to Production Order
                    oIssue.Lines.BaseEntry = prodOrderDocEntry;
                    oIssue.Lines.BaseLine = lineNum;
                    oIssue.Lines.BaseType = 202; // Production Order

                    //oIssue.Lines.ItemCode = itemCode;
                    //oIssue.Lines.WarehouseCode = whs;

                    double issueQty = 0;

                    // SERIAL-MANAGED ITEMS (Manual Issue Method)
                    //if (isSerial == "Y")
                    //{

                    //    string serialQuery = $@"
                    //    SELECT ""U_SerialType"" FROM ""OITM"" WHERE ""ItemCode"" = '{itemCode}'";

                    //    rsSer.DoQuery(serialQuery);

                    //    string serialType = rsSer.Fields.Item("U_SerialType").Value.ToString();
                    //    string mappedKeyValue = "";

                    //    // Map according to serial type
                    //    switch (serialType)
                    //    {
                    //        case "ENGINE": mappedKeyValue = rsMap.Fields.Item("U_EngineNo").Value.ToString(); break;
                    //        case "CHASSIS": mappedKeyValue = rsMap.Fields.Item("U_ChasisNo").Value.ToString(); break;
                    //        case "TRANSMISSION": mappedKeyValue = rsMap.Fields.Item("U_TransNo").Value.ToString(); break;
                    //        case "KEY": mappedKeyValue = rsMap.Fields.Item("U_SetKey").Value.ToString(); break;
                    //        default:
                    //            throw new Exception($"SerialType '{serialType}' not mapped for item {itemCode}");
                    //    }

                    //    // Always 1 
                    //    issueQty = 1;
                    //    oIssue.Lines.Quantity = issueQty;
                    //    // Attach serial
                    //    oIssue.Lines.SerialNumbers.InternalSerialNumber = mappedKeyValue;
                    //    oIssue.Lines.SerialNumbers.Quantity = issueQty;
                    //    oIssue.Lines.SerialNumbers.Add();
                    //}
                    //else
                    //{
                    //    issueQty = plannedQty/headerPlannedQty;
                    //    oIssue.Lines.Quantity = issueQty;
                    //}

                    if (rsBOM.Fields.Item("ManSerNum").Value.ToString() == "Y")
                    {
                        string serialType = rsBOM.Fields.Item("U_SerialType").Value.ToString();
                        string mappedValue = "";
                        if (serialType == "ENGINE") mappedValue = rsMap.Fields.Item("U_EngineNo").Value.ToString();
                        else if (serialType == "CHASSIS") mappedValue = rsMap.Fields.Item("U_ChasisNo").Value.ToString();
                        else if (serialType == "TRANSMISSION") mappedValue = rsMap.Fields.Item("U_TransNo").Value.ToString();
                        else if (serialType == "KEY") mappedValue = rsMap.Fields.Item("U_SetKey").Value.ToString();

                        oIssue.Lines.Quantity = 1;
                        oIssue.Lines.SerialNumbers.InternalSerialNumber = mappedValue;
                        oIssue.Lines.SerialNumbers.Quantity = 1;
                        oIssue.Lines.SerialNumbers.Add();
                    }
                    else
                    {
                        double pQty = Convert.ToDouble(rsBOM.Fields.Item("PlannedQty").Value);
                        double hQty = Convert.ToDouble(rsBOM.Fields.Item("headerQty").Value);
                        oIssue.Lines.Quantity = pQty / hQty;
                    }

                    oIssue.Lines.Add();
                    rsBOM.MoveNext();
                }
                //oProgBar.Value = 1; // Move bar to 50%
                //oProgBar.Text = "Posting Issue for Production... Please Wait.";

                // Add the Issue Document
                int ret = oIssue.Add();
                if (ret != 0)
                {
                    Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Issue creation failed: {errMsg}");
                }
                else
                {
                    //oProgBar.Value = 2;
                    string key = Utilities.Application.Company.GetNewObjectKey();
                    if (int.TryParse(key, out newDocEntry))
                    {
                        SBO_Application.StatusBar.SetText($"Auto Issue created (DocEntry: {newDocEntry})", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    }
                }

                // SBO_Application.StatusBar.SetText("Auto Issue for Production created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //SBO_Application.MessageBox("Auto Issue for Production created successfully.");
            }
            catch (Exception ex)
            {
                SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return 0;
            }

            finally
            {
                //    if (oProgBar != null)
                //    {
                //        oProgBar.Stop();
                //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgBar);
                //        oProgBar = null;
                //    }
                if (rsSer != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSer); rsSer = null; }
                if (rsMap != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsMap); rsMap = null; }
                if (rsBOM != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(rsBOM); rsBOM = null; }
                if (oIssue != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(oIssue); oIssue = null; }

                // Trigger Garbage Collection to ensure COM handles are fully dropped before SAP continues
                GC.Collect();
            }
            return newDocEntry;
        }
        #endregion

        #region Open Form for Scanning
        private void OpenUDOScanForm()
        {
            // 1. The Menu ID you provided
            string menuId = "2050";
            string formTypeToCheck = "139";

            try
            {
                SAPbouiCOM.Forms oForms = ((SAPbouiCOM.Application)this._Application).Forms;
                bool found = false;
                if (!found)
                {
                    ((SAPbouiCOM.Application)this._Application).ActivateMenuItem(menuId);
                    SAPbouiCOM.Form newForm = ((SAPbouiCOM.Application)this._Application).Forms.ActiveForm;
                    if (newForm.TypeEx == formTypeToCheck)
                    {
                        try
                        {
                            newForm.Items.Item("14").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            ((SAPbouiCOM.EditText)newForm.Items.Item("14").Specific).Active = true;
                        }
                        catch (Exception ex)
                        {

                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ((SAPbouiCOM.Application)this._Application).StatusBar.SetText(
                    "Error opening form via Menu: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion

        //private void OpenScanInBrowser(string scan)
        //{
        //    string html = $"<html><body style='font-family:Arial;font-size:22px;'>Scanned Code:<br><b>{scan}</b></body></html>";
        //    string url = "data:text/html," + Uri.EscapeDataString(html);

        //    Process.Start(new ProcessStartInfo
        //    {
        //        FileName = url,
        //        UseShellExecute = true
        //    });
        //}

        private void OpenScanInBrowser(string scan, string chasisNo = null, string station = null, string modelCode = null, string modelName = null)
        {
            bool showDetails =
                !string.IsNullOrEmpty(chasisNo) ||
                !string.IsNullOrEmpty(station) ||
                !string.IsNullOrEmpty(modelCode) ||
                !string.IsNullOrEmpty(modelName);

            string labelText = showDetails
                ? "Second scan taken for break down step:"
                : "Scan Chasis No:";

            string createdDate = DateTime.Now.ToString("dd-MMM-yyyy");
            string createdTime = DateTime.Now.ToString("HH:mm:ss");

            string detailsHtml = "";

            if (showDetails)
            {
                detailsHtml = $@"
                    <div class='details'>
                        <div><span>Chasis No</span>{chasisNo}</div>
                        <div><span>Model Name</span>{modelName}</div>
                        <div><span>Model Code</span>{modelCode}</div>
                        <div><span>Station</span>{station}</div>
                        <div><span>Created Date</span>{createdDate}</div>
                        <div><span>Created Time</span>{createdTime}</div>
                    </div>";
            }

            string html = $@"
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset='utf-8'>
            <title>Scan Result</title>

            <style>
            body {{
                margin: 0;
                height: 100vh;
                font-family: Segoe UI, Arial;
                background: radial-gradient(circle at top, #2196f3, #0d47a1);
                display: flex;
                align-items: center;
                justify-content: center;
            }}

            .container {{
                width: 900px;
                max-height: 90vh; /* ensures container never overflows viewport */
                background: white;
                border-radius: 24px;
                padding: 20px 35px;
                box-shadow: 0 25px 60px rgba(0,0,0,.35);
                text-align: center;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: flex-start;
                overflow: auto; /* allows scrolling if needed */
            }}

            .brand {{
                display: flex;
                flex-direction: column;
                align-items: center;
                margin-bottom: 15px;
            }}

            .logo svg {{
                height: 60px; /* slightly smaller to save space */
                fill: #0d47a1;
            }}

            .car-banner {{
                width: 105%;
                height: 100px;
                border-radius: 16px;
                margin-top: 10px;
                background: linear-gradient(135deg,#0d47a1,#42a5f5);
                display: flex;
                align-items: center;
                justify-content: center;
                color: white;
                font-size: 28px;
                font-weight: 700;
                letter-spacing: 2px;
            }}

            .label {{
                margin - top: 20px;
                font-size: 36px;
                font-weight: 700;
                color: #0d47a1;
            }}

            .scan {{
                margin - top: 10px;
                font-size: 50px;
                font-weight: 900;
                color: #0d47a1;
                background: linear-gradient(90deg,#bbdefb,#e3f2fd);
                padding: 10px 20px;
                border-radius: 16px;
                display: inline-block;
                word-break: break-all;
            }}

            .details {{
                margin - top: 20px;
                background: #e3f2fd;
                border-radius: 18px;
                padding: 15px 20px;
                text-align: left;
                font-size: 20px; /* slightly smaller font */
                width: 100%;
                box-sizing: border-box;
            }}

            .details div {{
                display: flex;
                justify-content: space-between;
                padding: 5px 0;
            }}

            .details span {{
                font - weight: 700;
                color: #1565c0;
            }}

            .countdown {{
                margin - top: 15px;
                font-size: 22px;
                font-weight: 700;
                color: #0d47a1;
            }};
            </style>

            <script>
            let sec = {_MESDBTime};
            setInterval(() => {{
                document.getElementById('cd').innerText =
                    'Closing in ' + sec + ' sec';
                sec--;
                if (sec < 0) window.close();
            }}, 1000);
            </script>

            </head>

            <body>
            <div class='container'>

                <div class='brand'>
                    <div class='car-banner'>
                        LAXMI MOTOR CORPS PRODUCTION LINE
                    </div>
                </div>

                <div class='label'>{labelText}</div>
                <div class='scan'>{scan}</div>

                {detailsHtml}

                <div id='cd' class='countdown'>Closing in {_MESDBTime} sec</div>

            </div>
            </body>
            </html>";

            string filePath = Path.Combine(Path.GetTempPath(), "ScanResult.html");
            File.WriteAllText(filePath, html);

            //Process.Start(new ProcessStartInfo
            //{
            //    FileName = filePath,
            //    UseShellExecute = true
            //});
            //System.Diagnostics.Process.Start("chrome.exe", filePath);
            System.Diagnostics.Process.Start("chrome.exe", "\"" + filePath + "\"");
        }


        private void CreateHiddenListenerForm()
        {
            string formUID = "frmBarScan";
            string itemUID = "txtScan";

            try
            {

                try
                {
                    SAPbouiCOM.Forms oForms = this._Application.Forms;
                    oForms.Item(formUID).Close();
                }
                catch { }

                SAPbouiCOM.FormCreationParams cp = (SAPbouiCOM.FormCreationParams)this._Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);

                cp.UniqueID = formUID;
                cp.FormType = "SCAN_BAR";
                cp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_FixedNoTitle;

                // 3. Add Form
                oScanForm = this._Application.Forms.AddEx(cp);
                oScanForm.AutoManaged = false;
                // 4. Set Properties (Off-screen)
                oScanForm.Left = 100000;
                oScanForm.Top = -100;
                oScanForm.ClientHeight = 1;
                oScanForm.ClientWidth = 1;
                oScanForm.Height = 1;
                oScanForm.Width = 1;

                oScanForm.Visible = true; // Must be true to receive keys

                // 5. Add EditText
                SAPbouiCOM.Item oItem = oScanForm.Items.Add(itemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oItem.Left = 0;
                oItem.Top = 0;
                oItem.Width = 1;
                oItem.Height = 1;
                oScanTxt = (SAPbouiCOM.EditText)oItem.Specific;
                // Success Log
                this._Application.StatusBar.SetText("Scanner Ready!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error creating scan form: " + ex.Message);
            }
        }
        private static void ForceScanFocus()
        {
            try
            {
                if (oScanForm != null)
                {
                    oScanForm.Select();
                    oScanForm.Items.Item("txtScan").Click();
                }
            }
            catch { }
        }
        private void ShowBrowserMessage(string icon, string subTitle, string chassis, string model, string modelName, string line, string time, string color = "#28a745", string nextAction = "", bool isUpdateOnly = false)
        {
            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus.html");

                // 1. Detect Error/Warning Keywords
                string stUpper = subTitle.ToUpper();
                bool isError = stUpper.Contains("INVALID ACTION") ||
                               stUpper.Contains("SEQUENCE ERROR") ||
                               stUpper.Contains("PAUSE/ STOP REQUIRED") ||
                               stUpper.Contains("ALREADY COMPLETED") ||
                               stUpper.Contains("START REQUIRED") ||
                               stUpper.Contains("RESUME REQUIRED") ||
                               stUpper.Contains("DATA NOT FOUND") ||
                               stUpper.Contains("INVALID ISSUE FOR PRODUCTION") || stUpper.Contains("INVALID CHASSIS");

                StringBuilder sidebarRows = new StringBuilder();
                string filterStatus = "";

                // 2. Identify current scanned status for the "Instant Injection" row
                if (stUpper.Contains("PAUSE")) filterStatus = "Pause";
                else if (stUpper.Contains("RESUME")) filterStatus = "Resume";
                else if (stUpper.Contains("FINISHED") || stUpper.Contains("STOP")) filterStatus = "Stop";
                else filterStatus = "Start";

                // 3. Status Board Logic
                if (!isError)
                {
                    // --- FIX: MANUALLY ADD THE CURRENT SCAN AT THE TOP ---
                    // This ensures the current scan shows up immediately even if DB update is lagging.
                    if (filterStatus != "Stop")
                    {
                        sidebarRows.Append($@"
                <div class='row' style='background-color: #fff9c4;'>
                    <div class='cell' style='font-weight:bold; color: #0d47a1;'>{chassis}</div>
                    <div class='cell status {filterStatus.ToLower()}'>{filterStatus}</div>
                </div>");
                    }

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // Exclude the current chassis from the query to prevent duplicate rows (since we injected it above)
                    string query = $@"
            SELECT ""U_ChassisNo"", ""U_Status""
            FROM ""@WJOBEXEH""
            WHERE ""U_JobDesc"" = '{line}' 
              AND ""U_Status"" <> 'Stop' 
              AND ""U_ChassisNo"" <> '{chassis}'
            ORDER BY ""U_StartTime"" DESC";

                    rs.DoQuery(query);
                    while (!rs.EoF)
                    {
                        string rowStatus = rs.Fields.Item("U_Status").Value.ToString();
                        sidebarRows.Append($@"
                <div class='row'>
                    <div class='cell' style='font-weight:bold;'>{rs.Fields.Item("U_ChassisNo").Value}</div>
                    <div class='cell status {rowStatus.ToLower()}'>{rowStatus}</div>
                </div>");
                        rs.MoveNext();
                    }
                }

                // Formatting Card Details (Left Side)
                string GetRow(string lbl, string val)
                {
                    if (string.IsNullOrEmpty(val) || val.ToUpper() == "N/A" || string.IsNullOrWhiteSpace(val)) return "";
                    return $@"<div class='card-row'><span>{lbl}</span> {val}</div>";
                }

                string detailsHtml = GetRow("Chassis No", chassis) + GetRow("Model Code", model) +
                                     GetRow("Model Name", modelName) + GetRow("Station/Line", line) + GetRow("Log Time", time);

                string actionHtml = string.IsNullOrEmpty(nextAction) ? "" :
                    $@"<div class='action-box'>
                <div class='action-title'>INSTRUCTION</div>
                <div class='action-text'>{nextAction}</div>
               </div>";

                // Generate Dynamic HTML with Height Matching
                string html = $@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ 
                    margin: 0; padding: 0; font-family: 'Segoe UI', Arial; 
                    background: #f1f3f5; height: 100vh; 
                    display: flex; align-items: center; justify-content: center; 
                    overflow: hidden; 
                }}
                .wrapper {{ 
                    display: flex; gap: 60px; align-items: stretch; 
                    justify-content: center; width: auto; max-width: 98%;
                }}
                .main-container {{ 
                    background: white; border-radius: 24px; padding: 30px; 
                    box-shadow: 0 15px 40px rgba(0,0,0,0.12); text-align: center; 
                    width: {(isError ? "900px" : "750px")}; 
                    flex-shrink: 0; border: 1px solid #e0e0e0;
                    display: flex; flex-direction: column; justify-content: center;
                }}
                .status-icon {{ font-size: 60px; color: {color}; margin-bottom: 5px; }}
                .sub-title {{ font-size: 28px; font-weight: 800; color: #333; text-transform: uppercase; margin-bottom: 15px; }}
                .details {{ background: #e3f2fd; border-radius: 15px; padding: 15px 20px; text-align: left; font-size: 20px; border-left: 10px solid {color}; }}
                .card-row {{ display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid rgba(0,0,0,0.05); }}
                .card-row span {{ font-weight: 700; color: #1565c0; }}
                .action-box {{ margin-top: 20px; background: #fffde7; border: 2px dashed #ffb300; border-radius: 12px; padding: 15px; text-align: left; }}
                .action-title {{ font-size: 13px; font-weight: 800; color: #f57f17; margin-bottom: 3px; }}
                .action-text {{ font-size: 18px; font-weight: 700; color: #333; }}

                .status-board {{ 
                    display: {(isError ? "none" : "flex")};
                    flex-direction: column; background: white; border-radius: 20px; 
                    width: 310px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); 
                    overflow: hidden; flex-shrink: 0; border: 1px solid #dee2e6;
                }}
                header {{ background: #0d47a1; color: white; padding: 12px; font-size: 13px; text-align: center; font-weight: bold; text-transform: uppercase; }}
                .grid-header, .row {{ display: flex; width: 100%; border-bottom: 1px solid #eee; }}
                .grid-header {{ background: #f8f9fa; font-weight: bold; color: #495057; font-size: 12px; }}
                .cell {{ flex: 1; padding: 10px; text-align: center; font-size: 13px; }}
                .scroll-box {{ flex: 1; overflow-y: auto; background: #fafafa; }}
                
                .status.start {{ color: #28a745; font-weight: bold; }}
                .status.pause {{ color: #fd7e14; font-weight: bold; }}
                .status.resume {{ color: #007bff; font-weight: bold; }}
                .status.stop {{ color: #dc3545; font-weight: bold; }}

                .footer-text {{ margin-top: 15px; font-size: 13px; font-weight: 600; color: #6c757d; }}
                .countdown {{ margin-top: 5px; font-size: 14px; font-weight: 700; color: #dc3545; }}
            </style>
            <script>
                let seconds = {_MESDBTime};
                setInterval(() => {{
                    seconds--;
                    document.getElementById('cd').innerText = 'Auto-closing in ' + seconds + ' seconds';
                    if (seconds <= 0) window.close();
                }}, 1000);
            </script>
        </head>
        <body>
            <div class='wrapper'>
                <div class='main-container'>
                    <div class='status-icon'>{icon}</div>
                    <div class='sub-title'>{subTitle}</div>
                    <div class='details'>{detailsHtml}</div>
                    {actionHtml}
                    <div class='footer-text'>LAXMI MOTOR CORPS</div>
                    <div id='cd' class='countdown'>Auto-closing in 5 seconds</div>
                </div>

                <div class='status-board'>
                    <header>{line} - ACTIVE JOBS</header>
                    <div class='grid-header'>
                        <div class='cell'>Chassis No</div>
                        <div class='cell'>Status</div>
                    </div>
                    <div class='scroll-box'>
                        {sidebarRows}
                    </div>
                </div>
            </div>
        </body>
        </html>";

                File.WriteAllText(filePath, html);
                if (!isUpdateOnly)
                {
                    //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    //System.Diagnostics.Process.Start("chrome.exe", filePath);
                    System.Diagnostics.Process.Start("chrome.exe", "\"" + filePath + "\"");
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Browser Error: " + ex.Message);
            }

        }

        private void ShowBrowserMessage1(string icon, string subTitle, string chassis, string model, string modelName, string line, string time, string color = "#28a745", string nextAction = "", bool isUpdateOnly = false)
        {
            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus1.html");

                // 1. Detect Error/Warning Keywords
                string stUpper = subTitle.ToUpper();
                bool isError = stUpper.Contains("INVALID ACTION") ||
                               stUpper.Contains("SEQUENCE ERROR") ||
                               stUpper.Contains("PAUSE/ STOP REQUIRED") ||
                               stUpper.Contains("ALREADY COMPLETED") ||
                               stUpper.Contains("START REQUIRED") ||
                               stUpper.Contains("RESUME REQUIRED") ||
                               stUpper.Contains("DATA NOT FOUND") ||
                               stUpper.Contains("INVALID ISSUE FOR PRODUCTION") || stUpper.Contains("INVALID CHASSIS");

                StringBuilder sidebarRows = new StringBuilder();
                string filterStatus = "";

                // 2. Identify current scanned status for the "Instant Injection" row
                if (stUpper.Contains("PAUSE")) filterStatus = "Pause";
                else if (stUpper.Contains("RESUME")) filterStatus = "Resume";
                else if (stUpper.Contains("FINISHED") || stUpper.Contains("STOP")) filterStatus = "Stop";
                else filterStatus = "Start";

                // 3. Status Board Logic
                if (!isError)
                {
                    // --- FIX: MANUALLY ADD THE CURRENT SCAN AT THE TOP ---
                    // This ensures the current scan shows up immediately even if DB update is lagging.
                    if (filterStatus != "Stop")
                    {
                        sidebarRows.Append($@"
                <div class='row' style='background-color: #fff9c4;'>
                    <div class='cell' style='font-weight:bold; color: #0d47a1;'>{chassis}</div>
                    <div class='cell status {filterStatus.ToLower()}'>{filterStatus}</div>
                </div>");
                    }

                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    // Exclude the current chassis from the query to prevent duplicate rows (since we injected it above)
                    string query = $@"
            SELECT ""U_ChassisNo"", ""U_Status""
            FROM ""@WJOBEXEH""
            WHERE ""U_JobDesc"" = '{line}' 
              AND ""U_Status"" <> 'Stop' 
              AND ""U_ChassisNo"" <> '{chassis}'
            ORDER BY ""U_StartTime"" DESC";

                    rs.DoQuery(query);
                    while (!rs.EoF)
                    {
                        string rowStatus = rs.Fields.Item("U_Status").Value.ToString();
                        sidebarRows.Append($@"
                <div class='row'>
                    <div class='cell' style='font-weight:bold;'>{rs.Fields.Item("U_ChassisNo").Value}</div>
                    <div class='cell status {rowStatus.ToLower()}'>{rowStatus}</div>
                </div>");
                        rs.MoveNext();
                    }
                }

                // Formatting Card Details (Left Side)
                string GetRow(string lbl, string val)
                {
                    if (string.IsNullOrEmpty(val) || val.ToUpper() == "N/A" || string.IsNullOrWhiteSpace(val)) return "";
                    return $@"<div class='card-row'><span>{lbl}</span> {val}</div>";
                }

                string detailsHtml = GetRow("Chassis No", chassis) + GetRow("Model Code", model) +
                                     GetRow("Model Name", modelName) + GetRow("Station/Line", line) + GetRow("Log Time", time);

                string actionHtml = string.IsNullOrEmpty(nextAction) ? "" :
                    $@"<div class='action-box'>
                <div class='action-title'>INSTRUCTION</div>
                <div class='action-text'>{nextAction}</div>
               </div>";

                // Generate Dynamic HTML with Height Matching
                string html = $@"
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset='utf-8'>
            <style>
                body {{ 
                    margin: 0; padding: 0; font-family: 'Segoe UI', Arial; 
                    background: #f1f3f5; height: 100vh; 
                    display: flex; align-items: center; justify-content: center; 
                    overflow: hidden; 
                }}
                .wrapper {{ 
                    display: flex; gap: 60px; align-items: stretch; 
                    justify-content: center; width: auto; max-width: 98%;
                }}
                .main-container {{ 
                    background: white; border-radius: 24px; padding: 30px; 
                    box-shadow: 0 15px 40px rgba(0,0,0,0.12); text-align: center; 
                    width: {(isError ? "900px" : "750px")}; 
                    flex-shrink: 0; border: 1px solid #e0e0e0;
                    display: flex; flex-direction: column; justify-content: center;
                }}
                .status-icon {{ font-size: 60px; color: {color}; margin-bottom: 5px; }}
                .sub-title {{ font-size: 28px; font-weight: 800; color: #333; text-transform: uppercase; margin-bottom: 15px; }}
                .details {{ background: #e3f2fd; border-radius: 15px; padding: 15px 20px; text-align: left; font-size: 20px; border-left: 10px solid {color}; }}
                .card-row {{ display: flex; justify-content: space-between; padding: 6px 0; border-bottom: 1px solid rgba(0,0,0,0.05); }}
                .card-row span {{ font-weight: 700; color: #1565c0; }}
                .action-box {{ margin-top: 20px; background: #fffde7; border: 2px dashed #ffb300; border-radius: 12px; padding: 15px; text-align: left; }}
                .action-title {{ font-size: 13px; font-weight: 800; color: #f57f17; margin-bottom: 3px; }}
                .action-text {{ font-size: 18px; font-weight: 700; color: #333; }}

                .status-board {{ 
                    display: {(isError ? "none" : "flex")};
                    flex-direction: column; background: white; border-radius: 20px; 
                    width: 310px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); 
                    overflow: hidden; flex-shrink: 0; border: 1px solid #dee2e6;
                }}
                header {{ background: #0d47a1; color: white; padding: 12px; font-size: 13px; text-align: center; font-weight: bold; text-transform: uppercase; }}
                .grid-header, .row {{ display: flex; width: 100%; border-bottom: 1px solid #eee; }}
                .grid-header {{ background: #f8f9fa; font-weight: bold; color: #495057; font-size: 12px; }}
                .cell {{ flex: 1; padding: 10px; text-align: center; font-size: 13px; }}
                .scroll-box {{ flex: 1; overflow-y: auto; background: #fafafa; }}
                
                .status.start {{ color: #28a745; font-weight: bold; }}
                .status.pause {{ color: #fd7e14; font-weight: bold; }}
                .status.resume {{ color: #007bff; font-weight: bold; }}
                .status.stop {{ color: #dc3545; font-weight: bold; }}

                .footer-text {{ margin-top: 15px; font-size: 13px; font-weight: 600; color: #6c757d; }}
                .countdown {{ margin-top: 5px; font-size: 14px; font-weight: 700; color: #dc3545; }}
            </style>
            <script>
                let seconds = {_MESDBTime};
                setInterval(() => {{
                    seconds--;
                    document.getElementById('cd').innerText = 'Auto-closing in ' + seconds + ' seconds';
                    if (seconds <= 0) window.close();
                }}, 1000);
            </script>
        </head>
        <body>
            <div class='wrapper'>
                <div class='main-container'>
                    <div class='status-icon'>{icon}</div>
                    <div class='sub-title'>{subTitle}</div>
                    <div class='details'>{detailsHtml}</div>
                    {actionHtml}
                    <div class='footer-text'>LAXMI MOTOR CORPS</div>
                    <div id='cd' class='countdown'>Auto-closing in 5 seconds</div>
                </div>

                <div class='status-board'>
                    <header>{line} - ACTIVE JOBS</header>
                    <div class='grid-header'>
                        <div class='cell'>Chassis No</div>
                        <div class='cell'>Status</div>
                    </div>
                    <div class='scroll-box'>
                        {sidebarRows}
                    </div>
                </div>
            </div>
        </body>
        </html>";

                File.WriteAllText(filePath, html);
                if (!isUpdateOnly)
                {
                    //System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    //System.Diagnostics.Process.Start("chrome.exe", filePath);
                    System.Diagnostics.Process.Start("chrome.exe", "\"" + filePath + "\"");
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Browser Error: " + ex.Message);
            }

        }

        //private void ShowBrowserMessage(string icon, string subTitle, string chassis, string model, string modelName, string line, string time, string color = "#28a745", string nextAction = "")
        //{
        //    try
        //    {
        //        string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus.html");

        //        // Helper to generate a row only if value is not empty
        //        string GetRow(string label, string value)
        //        {
        //            if (string.IsNullOrEmpty(value) || value.ToUpper() == "N/A") return "";
        //            return $@"<div><span>{label}</span> {value}</div>";
        //        }

        //        // Build the dynamic list of details
        //        string detailsHtml = "";
        //        detailsHtml += GetRow("Chassis No", chassis);
        //        detailsHtml += GetRow("Model Code", model);
        //        detailsHtml += GetRow("Model Name", modelName);
        //        detailsHtml += GetRow("Station/Line", line);
        //        detailsHtml += GetRow("Log Time", time);

        //        // Build the instruction box only if nextAction is provided
        //        // Styled to match your light theme (using a very light version of the status color)
        //        string actionHtml = string.IsNullOrEmpty(nextAction) ? "" :
        //            $@"<div class='action-box'>
        //        <div class='action-title'>NEXT ACTION</div>
        //        <div class='action-text'>{nextAction}</div>
        //       </div>";

        //        string html = $@"
        //<!DOCTYPE html>
        //<html>
        //<head>
        //    <meta charset='utf-8'>
        //    <title>Scanned Job Status</title>
        //    <style>
        //        body {{
        //            margin: 0;
        //            height: 100vh;
        //            font-family: 'Segoe UI', Arial;
        //            background: #f4f7f6; /* Your preferred background */
        //            display: flex;
        //            align-items: center;
        //            justify-content: center;
        //            overflow: hidden;
        //        }}

        //        .container {{
        //            width: 850px;
        //            background: white;
        //            border-radius: 24px;
        //            padding: 30px 40px;
        //            box-shadow: 0 25px 60px rgba(0,0,0,.35);
        //            text-align: center;
        //            display: flex;
        //            flex-direction: column;
        //            align-items: center;
        //        }}

        //        .status-icon {{
        //            font-size: 55px; /* Slightly larger for visibility */
        //            font-weight: 900;
        //            color: {color};
        //            margin-bottom: 5px;
        //        }}

        //        .sub-title {{
        //            font-size: 36px;
        //            font-weight: 700;
        //            color: #333;
        //            text-transform: uppercase;
        //            margin-bottom: 25px;
        //        }}

        //        .details {{
        //            background: #e3f2fd; /* Your preferred light blue */
        //            border-radius: 18px;
        //            padding: 20px 30px;
        //            text-align: left;
        //            font-size: 24px;
        //            width: 100%;
        //            box-sizing: border-box;
        //            border-left: 10px solid {color};
        //        }}

        //        .details div {{
        //            display: flex;
        //            justify-content: space-between;
        //            padding: 10px 0;
        //            border-bottom: 1px solid rgba(0,0,0,0.05);
        //        }}

        //        .details div:last-child {{ border-bottom: none; }}

        //        .details span {{
        //            font-weight: 700;
        //            color: #1565c0; /* Your preferred blue value color */
        //        }}

        //        /* Instruction/Action Box */
        //        .action-box {{
        //            margin-top: 20px;
        //            background: #fff;
        //            border: 2px dashed {color};
        //            border-radius: 15px;
        //            padding: 15px 25px;
        //            width: 100%;
        //            box-sizing: border-box;
        //            text-align: left;
        //        }}
        //        .action-title {{ font-size: 16px; font-weight: 800; color: {color}; margin-bottom: 5px; }}
        //        .action-text {{ font-size: 24px; font-weight: 700; color: #333; }}

        //        .footer-text {{
        //            margin-top: 25px;
        //            font-size: 16px;
        //            font-weight: 600;
        //            color: #0d47a1;
        //            opacity: 0.8;
        //        }}

        //        .countdown {{
        //            margin-top: 10px;
        //            font-size: 18px;
        //            font-weight: 700;
        //            color: #d32f2f;
        //        }}
        //    </style>
        //    <script>
        //        let seconds = 10;
        //        function startCountdown() {{
        //            const cd = document.getElementById('cd');
        //            const timer = setInterval(() => {{
        //                seconds--;
        //                cd.innerText = 'Auto-closing in ' + seconds + ' seconds';
        //                if (seconds <= 0) {{
        //                    clearInterval(timer);
        //                    window.close();
        //                }}
        //            }}, 1000);
        //        }}
        //    </script>
        //</head>

        //<body onload='startCountdown()'>
        //    <div class='container'>
        //        <div class='status-icon'>{icon}</div>
        //        <div class='sub-title'>{subTitle}</div>

        //        <div class='details'>
        //            {detailsHtml}
        //        </div>

        //        {actionHtml}

        //        <div class='footer-text'>LAXMI MOTOR CORPS - PRODUCTION LINE</div>
        //        <div id='cd' class='countdown'>Auto-closing in 10 seconds</div>
        //    </div>
        //</body>
        //</html>";

        //        // 3. Save and Open
        //        File.WriteAllText(filePath, html);

        //        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
        //        {
        //            FileName = filePath,
        //            UseShellExecute = true
        //        });
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.SetStatusBarMessage("Browser Error: " + ex.Message);
        //    }
        //}

        private void ShowProcessingMessage(string chassis, string model, string modelName, string line)
        {
            try
            {
                string filePath = Path.Combine(Path.GetTempPath(), "ScannedJobStatus.html");
                string time = DateTime.Now.ToString("HH:mm:ss");

                string html = $@"
                <!DOCTYPE html>
                <html>
                <head>
                    <meta charset='utf-8'>
                    <style>
                        body {{ margin: 0; padding: 0; font-family: 'Segoe UI'; background: #f1f3f5; display: flex; align-items: center; justify-content: center; height: 100vh; overflow: hidden; }}
                        .main-container {{ 
                            background: white; border-radius: 24px; padding: 40px; 
                            box-shadow: 0 15px 40px rgba(0,0,0,0.15); width: 700px; 
                            text-align: center; border-top: 15px solid #FFC107; 
                        }}
                        .spinner {{
                            width: 80px; height: 80px; border: 8px solid #f3f3f3; border-top: 8px solid #FFC107;
                            border-radius: 50%; animation: spin 1s linear infinite; margin: 0 auto 20px;
                        }}
                        @keyframes spin {{ 0% {{ transform: rotate(0deg); }} 100% {{ transform: rotate(360deg); }} }}
                        .sub-title {{ font-size: 28px; font-weight: 800; color: #333; margin-bottom: 10px; }}
                        .details {{ background: #fffde7; border-radius: 15px; padding: 20px; text-align: left; font-size: 22px; border: 1px solid #ffe082; }}
                        .card-row {{ display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px solid rgba(0,0,0,0.05); }}
                        .card-row span {{ font-weight: 700; color: #f57f17; }}
                        .wait-msg {{ margin-top: 25px; font-size: 18px; font-weight: 700; color: #666; }}
                        #cd {{ color: #d32f2f; font-weight: bold; margin-top: 10px; font-size: 14px; }}
                    </style>
                </head>
                <body>
                    <div class='main-container'>
                        <div class='spinner'></div>
                        <div class='sub-title'>PROCESSING ISSUE...</div>
                        <div class='details'>
                            <div class='card-row'><span>Chassis No</span> {chassis}</div>
                            <div class='card-row'><span>Model</span> {modelName}</div>
                            <div class='card-row'><span>Station</span> {line}</div>
                            <div class='card-row'><span>Started At</span> {time}</div>
                        </div>
                        <div class='wait-msg'>⚠️ Creating Issue For Production.</div>
                        <!-- ADDED THE MISSING ELEMENT HERE -->
                        <div id='cd'>Auto-closing in {_MESDBTime} seconds</div>
                    </div>

                    <script>
                        let seconds = {_MESDBTime};
                        let countdown = setInterval(() => {{
                            seconds--;
                            let el = document.getElementById('cd');
            
                            if (seconds <= 0) {{
                                clearInterval(countdown);
                                // Attempt to close
                                window.close();
                                // Fallback for Chrome: redirect to a blank page if close is blocked
                                setTimeout(() => {{
                                    window.location.href = 'about:blank';
                                }}, 1000);
                            }}
                        }}, 1000);
                    </script>
                </body>
                </html>";

                File.WriteAllText(filePath, html);
                System.Diagnostics.Process.Start("chrome.exe", "\"" + filePath + "\"");
            }
            catch (Exception ex) { }
        }

        //private void OpenTrimAssemblyStatusBoard()
        //{
        //    SAPbobsCOM.Recordset rs = null;

        //    try
        //    {
        //        if (_Company == null || !_Company.Connected)
        //            return;

        //        rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        string query = @"
        //    SELECT T0.""U_ChassisNo"", T0.""U_FGDesc"", T0.""U_JobDesc"", T0.""U_JobID"", 
        //           T0.""U_Status"", T0.""U_StartDate"", T0.""U_StartTime""
        //    FROM ""@WJOBEXEH"" T0
        //    WHERE T0.""U_Status"" <> 'Stop'
        //    ORDER BY T0.""U_StartDate"" DESC, T0.""U_StartTime"" DESC";

        //        rs.DoQuery(query);

        //        StringBuilder rows = new StringBuilder();

        //        while (!rs.EoF)
        //        {
        //            string startDate = rs.Fields.Item("U_StartDate").Value != null ?
        //                Convert.ToDateTime(rs.Fields.Item("U_StartDate").Value).ToString("dd MMM yyyy") : "-";

        //            string startTime = "-";
        //            if (rs.Fields.Item("U_StartTime").Value != null)
        //            {
        //                int t = Convert.ToInt32(rs.Fields.Item("U_StartTime").Value);
        //                startTime = DateTime.Today.AddHours(t / 100).AddMinutes(t % 100).ToString("hh:mm tt");
        //            }

        //            string status = rs.Fields.Item("U_Status").Value.ToString();
        //            string statusClass = status.ToLower();

        //            rows.Append($@"
        //        <div class='row'>
        //            <div class='cell chassis'>{rs.Fields.Item("U_ChassisNo").Value}</div>
        //            <div class='cell product'>{rs.Fields.Item("U_FGDesc").Value}</div>
        //            <div class='cell'>{rs.Fields.Item("U_JobDesc").Value}</div>
        //            <div class='cell jobid'>{rs.Fields.Item("U_JobID").Value}</div>
        //            <div class='cell'>{startDate}</div>
        //            <div class='cell'>{startTime}</div>
        //            <div class='cell'><span class='badge {statusClass}'>{status}</span></div>
        //        </div>");

        //            rs.MoveNext();
        //        }

        //        string html = $@"<!DOCTYPE html>
        //<html>
        //<head>
        //    <meta charset='utf-8'>
        //    <title>LMC Status Board</title>
        //    <style>
        //        /* DARK MODE (Default) */
        //        :root {{
        //            --bg: #0f172a;
        //            --card: #1e293b;
        //            --header-bg: #1e293b;
        //            --grid-head: #334155;
        //            --text: #f8fafc;
        //            --text-dim: #94a3b8;
        //            --accent: #38bdf8;
        //            --border: #334155;
        //            --success: #22c55e;
        //            --warning: #f59e0b;
        //        }}

        //        /* LIGHT MODE OVERRIDES */
        //        body.light-mode {{
        //            --bg: #f1f5f9;
        //            --card: #ffffff;
        //            --header-bg: #ffffff;
        //            --grid-head: #e2e8f0;
        //            --text: #0f172a;
        //            --text-dim: #64748b;
        //            --accent: #0284c7;
        //            --border: #e2e8f0;
        //            --success: #16a34a;
        //            --warning: #d97706;
        //        }}

        //        body {{
        //            margin: 0; padding: 0;
        //            font-family: 'Segoe UI', system-ui, sans-serif;
        //            background-color: var(--bg);
        //            color: var(--text);
        //            transition: background-color 0.3s, color 0.3s;
        //            overflow: hidden;
        //        }}

        //        header {{
        //            background: var(--header-bg);
        //            padding: 15px 40px;
        //            display: flex;
        //            justify-content: space-between;
        //            align-items: center;
        //            border-bottom: 3px solid var(--accent);
        //            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        //        }}

        //        .header-left {{ display: flex; align-items: center; gap: 20px; }}

        //        .dot {{
        //            width: 12px; height: 12px; background: #ef4444;
        //            border-radius: 50%; animation: pulse 1.5s infinite;
        //        }}
        //        @keyframes pulse {{
        //            0% {{ transform: scale(0.9); opacity: 1; }}
        //            50% {{ transform: scale(1.3); opacity: 0.5; }}
        //            100% {{ transform: scale(0.9); opacity: 1; }}
        //        }}

        //        .title {{ font-size: 22px; font-weight: 800; letter-spacing: 1px; }}

        //        .controls {{ display: flex; align-items: center; gap: 25px; }}

        //        .clock {{ font-size: 18px; font-family: monospace; color: var(--accent); font-weight: bold; }}

        //        #themeBtn {{
        //            background: var(--accent);
        //            color: white;
        //            border: none;
        //            padding: 8px 16px;
        //            border-radius: 6px;
        //            cursor: pointer;
        //            font-weight: bold;
        //            font-size: 12px;
        //            text-transform: uppercase;
        //            transition: opacity 0.2s;
        //        }}
        //        #themeBtn:hover {{ opacity: 0.8; }}

        //        .container {{ padding: 20px; }}

        //        .grid-header {{
        //            display: flex; background: var(--grid-head);
        //            padding: 15px; border-radius: 8px 8px 0 0;
        //            font-weight: bold; font-size: 12px; color: var(--text-dim);
        //            text-transform: uppercase; border: 1px solid var(--border);
        //        }}

        //        .scroll-box {{
        //            height: calc(100vh - 165px);
        //            overflow-y: auto;
        //            background: var(--card);
        //            border-radius: 0 0 8px 8px;
        //            border: 1px solid var(--border);
        //            border-top: none;
        //        }}

        //        .row {{
        //            display: flex; align-items: center;
        //            padding: 15px; border-bottom: 1px solid var(--border);
        //            transition: background 0.2s;
        //        }}
        //        .row:hover {{ background: rgba(0,0,0,0.02); }}

        //        .cell {{ flex: 1; padding: 0 10px; text-align: center; }}
        //        .chassis {{ font-weight: bold; color: var(--accent); font-size: 17px; font-family: monospace; }}
        //        .product {{ flex: 1.5; text-align: left; font-weight: 500; }}
        //        .jobid {{ font-size: 12px; color: var(--text-dim); }}

        //        .badge {{
        //            padding: 4px 12px; border-radius: 12px; font-size: 11px; font-weight: 800;
        //            border: 1px solid; text-transform: uppercase;
        //        }}
        //        .badge.start {{ border-color: var(--success); color: var(--success); background: rgba(34,197,94,0.1); }}
        //        .badge.pause {{ border-color: var(--warning); color: var(--warning); background: rgba(245,158,11,0.1); }}
        //        .badge.resume {{ border-color: var(--accent); color: var(--accent); background: rgba(56,189,248,0.1); }}

        //        .scroll-box::-webkit-scrollbar {{ width: 0px; }}
        //    </style>

        //    <script>
        //        function toggleTheme() {{
        //            const body = document.body;
        //            body.classList.toggle('light-mode');
        //            const isLight = body.classList.contains('light-mode');
        //            document.getElementById('themeBtn').innerText = isLight ? '🌙 DARK MODE' : '☀️ LIGHT MODE';
        //            localStorage.setItem('theme', isLight ? 'light' : 'dark');
        //        }}

        //        function updateClock() {{
        //            document.getElementById('clock').innerText = new Date().toLocaleTimeString();
        //        }}

        //        function autoScroll() {{
        //            const box = document.getElementById('scrollBox');
        //            setInterval(() => {{
        //                box.scrollTop += 1;
        //                if (box.scrollTop + box.clientHeight >= box.scrollHeight - 1) {{
        //                    setTimeout(() => {{ box.scrollTop = 0; }}, 3000);
        //                }}
        //            }}, 50);
        //        }}

        //        window.onload = () => {{
        //            // Load saved theme
        //            if (localStorage.getItem('theme') === 'light') toggleTheme();

        //            updateClock();
        //            setInterval(updateClock, 1000);
        //            autoScroll();
        //        }};
        //    </script>
        //</head>
        //<body>
        //    <header>
        //        <div class='header-left'>
        //            <div class='dot'></div>
        //            <div class='title'>TRIM ASSEMBLY <span style='font-weight:300;'>LINE STATUS</span></div>
        //        </div>
        //        <div class='controls'>
        //            <div id='clock' class='clock'>00:00:00</div>
        //            <button id='themeBtn' onclick='toggleTheme()'>☀️ LIGHT MODE</button>
        //        </div>
        //    </header>

        //    <div class='container'>
        //        <div class='grid-header'>
        //            <div class='cell'>Chassis</div>
        //            <div class='cell' style='flex:1.5; text-align:left;'>Product Description</div>
        //            <div class='cell'>Operation</div>
        //            <div class='cell'>Job ID</div>
        //            <div class='cell'>Date</div>
        //            <div class='cell'>Time</div>
        //            <div class='cell'>Status</div>
        //        </div>
        //        <div id='scrollBox' class='scroll-box'>
        //            {rows}
        //        </div>
        //    </div>
        //</body>
        //</html>";

        //        string filePath = Path.Combine(Path.GetTempPath(), "TrimAssemblyStatusBoard.html");
        //        File.WriteAllText(filePath, html);
        //        Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { { SBO_Application.MessageBox(ex.Message); } }
        //    finally { { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); } }
        //}

        // with each tabs trim, chasis, final and test
        //private void OpenTrimAssemblyStatusBoard()
        //{
        //    SAPbobsCOM.Recordset rs = null;
        //    try
        //    {
        //        if (_Company == null || !_Company.Connected) return;
        //        rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        string query = @"
        //            SELECT T0.""U_ChassisNo"", T0.""U_FGDesc"", T0.""U_JobDesc"", T0.""U_JobID"", 
        //                   T0.""U_Status"", T0.""U_StartDate"", T0.""U_StartTime""
        //            FROM ""@WJOBEXEH"" T0
        //            WHERE T0.""U_Status"" <> 'Stop'
        //            ORDER BY T0.""U_JobDesc"" ASC, T0.""U_StartDate"" DESC";

        //        rs.DoQuery(query);

        //        // Group rows by Operation Name
        //        var opGroups = new System.Collections.Generic.Dictionary<string, StringBuilder>();

        //        while (!rs.EoF)
        //        {
        //            string opName = rs.Fields.Item("U_JobDesc").Value.ToString();
        //            if (string.IsNullOrEmpty(opName)) opName = "General Assembly";

        //            if (!opGroups.ContainsKey(opName)) opGroups.Add(opName, new StringBuilder());

        //            string startDate = rs.Fields.Item("U_StartDate").Value != null ?
        //                Convert.ToDateTime(rs.Fields.Item("U_StartDate").Value).ToString("dd MMM") : "-";

        //            string status = rs.Fields.Item("U_Status").Value.ToString();
        //            string statusClass = status.ToLower();

        //            opGroups[opName].Append($@"
        //                <div class='row'>
        //                    <div class='cell chassis'>{rs.Fields.Item("U_ChassisNo").Value}</div>
        //                    <div class='cell product' title='{rs.Fields.Item("U_FGDesc").Value}'>{rs.Fields.Item("U_FGDesc").Value}</div>
        //                    <div class='cell status'><span class='badge {statusClass}'>{status}</span></div>
        //                    <div class='cell date'>{startDate}</div>
        //                </div>");
        //            rs.MoveNext();
        //        }

        //        // --- CUSTOM SORTING LOGIC ---
        //        // Define the desired sequence
        //        var sortedKeys = new System.Collections.Generic.List<string>(opGroups.Keys);
        //        sortedKeys.Sort((x, y) =>
        //        {
        //            int GetPriority(string name)
        //            {
        //                string n = name.ToUpper();
        //                if (n.Contains("TRIM")) return 1;
        //                if (n.Contains("CHASIS")) return 2;
        //                if (n.Contains("FINAL")) return 3;
        //                if (n.Contains("TEST")) return 4;
        //                return 5; // Everything else
        //            }
        //            return GetPriority(x).CompareTo(GetPriority(y));
        //        });

        //        StringBuilder boxesHtml = new StringBuilder();
        //        foreach (string key in sortedKeys)
        //        {
        //            boxesHtml.Append($@"
        //                <div class='op-card'>
        //                    <div class='op-header'>{key}</div>
        //                    <div class='grid-header'>
        //                        <div class='col col-chassis'>CHASSIS</div>
        //                        <div class='col col-product'>PRODUCT</div>
        //                        <div class='col col-status'>STATUS</div>
        //                        <div class='col col-date'>DATE</div>
        //                    </div>
        //                    <div class='scroll-box'>
        //                        {opGroups[key].ToString()}
        //                    </div>
        //                </div>");
        //        }

        //        // The CSS remains the same high-quality visual style with "cutting" fixes
        //        string html = $@"<!DOCTYPE html>
        //<html>
        //<head>
        //    <meta charset='utf-8'>
        //    <style>
        //        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');

        //        :root {{
        //            --bg: #0b0e14;
        //            --card-bg: rgba(30, 41, 59, 0.7);
        //            --accent: #38bdf8;
        //            --success: #22c55e;
        //            --warning: #eab308;
        //            --text-main: #f8fafc;
        //            --text-dim: #94a3b8;
        //        }}

        //        body {{ 
        //            margin: 0; padding: 40px; font-family: 'Inter', sans-serif; 
        //            background: radial-gradient(circle at top right, #1e293b, #0b0e14);
        //            color: var(--text-main); min-height: 100vh;
        //        }}

        //        h1 {{ text-align: center; font-weight: 800; font-size: 42px; margin-bottom: 50px; text-transform: uppercase; letter-spacing: -1px; background: linear-gradient(to bottom, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}

        //        .dashboard-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(650px, 1fr)); gap: 30px; }}

        //        .op-card {{ background: var(--card-bg); backdrop-filter: blur(12px); border-radius: 20px; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); display: flex; flex-direction: column; }}

        //        .op-header {{ padding: 20px 25px; font-size: 22px; font-weight: 800; color: var(--accent); border-bottom: 1px solid rgba(255,255,255,0.05); }}

        //        .grid-header {{ display: flex; padding: 12px 25px; background: rgba(0,0,0,0.2); font-size: 11px; font-weight: 700; color: var(--text-dim); letter-spacing: 1px; }}

        //        .scroll-box {{ max-height: 450px; overflow-y: auto; scrollbar-width: none; padding: 10px 0; }}
        //        .scroll-box::-webkit-scrollbar {{ display: none; }}

        //        .row {{ display: flex; align-items: center; padding: 15px 25px; border-bottom: 1px solid rgba(255,255,255,0.03); }}
        //        .row:hover {{ background: rgba(255,255,255,0.03); }}

        //        .cell {{ padding: 0 5px; overflow: hidden; }}
        //        .col-chassis, .chassis {{ width: 30%; font-family: monospace; font-weight: 600; color: var(--accent); }}
        //        .col-product, .product {{ width: 40%; font-size: 13px; color: #cbd5e1; }}
        //        .col-status, .status {{ width: 15%; text-align: center; }}
        //        .col-date, .date {{ width: 15%; text-align: right; color: var(--text-dim); font-size: 12px; }}

        //        .product {{ display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; line-height: 1.4; white-space: normal; }}

        //        .badge {{ padding: 4px 10px; border-radius: 8px; font-size: 10px; font-weight: 800; display: inline-block; width: 65px; text-align: center; }}
        //        .badge.start {{ background: rgba(34,197,94,0.15); color: var(--success); border: 1px solid var(--success); }}
        //        .badge.pause {{ background: rgba(234,179,8,0.15); color: var(--warning); border: 1px solid var(--warning); }}
        //        .badge.resume {{ background: rgba(56,189,248,0.15); color: var(--accent); border: 1px solid var(--accent); }}
        //    </style>

        //    <script>
        //        function autoScroll() {{
        //            const boxes = document.querySelectorAll('.scroll-box');
        //            setInterval(() => {{
        //                boxes.forEach(box => {{
        //                    if (box.scrollHeight > box.clientHeight) {{
        //                        box.scrollTop += 1;
        //                        if (box.scrollTop + box.clientHeight >= box.scrollHeight - 1) {{
        //                            setTimeout(() => {{ box.scrollTop = 0; }}, 2000);
        //                        }}
        //                    }}
        //                }});
        //            }}, 50);
        //        }}
        //        window.onload = autoScroll;
        //    </script>
        //</head>
        //<body>
        //    <h1>LMC Production Dashboard</h1>
        //    <div class='dashboard-grid'>
        //        {boxesHtml}
        //    </div>
        //</body>
        //</html>";

        //        string filePath = Path.Combine(Path.GetTempPath(), "LMC_Dashboard.html");
        //        File.WriteAllText(filePath, html);
        //        Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
        //    }
        //    catch (Exception ex) { SBO_Application.MessageBox(ex.Message); }
        //    finally { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        //}

        private void OpenTrimAssemblyStatusBoard1()
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                if (_Company == null || !_Company.Connected) return;
                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // --- SECTION 1: ACTIVE JOBS ---
                //    string queryActive = @"
                //SELECT T0.""U_ChassisNo"", T0.""U_FGDesc"", T0.""U_JobDesc"", T0.""U_JobID"", 
                //       T0.""U_Status"", T0.""U_StartDate"", T0.""U_StartTime""
                //FROM ""@WJOBEXEH"" T0
                //WHERE T0.""U_Status"" <> 'Stop'
                //ORDER BY T0.""U_JobDesc"" ASC, T0.""U_StartDate"" DESC";
                string queryActive = @"
                    SELECT
                    T0.""U_ChassisNo"", 
                    T0.""U_FGDesc"", 
                    T0.""U_JobDesc"", 
                    T0.""U_Status"", 
                    T0.""U_StartTime"",
                    CASE
                        WHEN T0.""U_Status"" = 'Start' THEN T1.""U_FirstBarcode""
                        WHEN T0.""U_Status"" = 'Resume' THEN T1.""U_JobResume""
                        WHEN T0.""U_Status"" = 'Pause' THEN T1.""U_JobPause""
                        ELSE T1.""U_LastBarcode""
                    END AS ""Current_Barcode""
                FROM ""@WJOBEXEH"" T0
                LEFT JOIN ""@WRKORDRDTLSC"" T1 ON T1.""U_WorkId"" = T0.""U_JobID""
                WHERE T0.""U_Status"" <> 'Stop'
                ORDER BY T0.""U_StartTime"" DESC";

                rs.DoQuery(queryActive);
                var opGroups = new System.Collections.Generic.Dictionary<string, StringBuilder>();

                while (!rs.EoF)
                {
                    string opName = rs.Fields.Item("U_JobDesc").Value.ToString();
                    if (string.IsNullOrEmpty(opName)) opName = "General Assembly";
                    if (!opGroups.ContainsKey(opName)) opGroups.Add(opName, new StringBuilder());

                    string startDate = rs.Fields.Item("U_StartDate").Value != null ?
                        Convert.ToDateTime(rs.Fields.Item("U_StartDate").Value).ToString("dd MMM") : "-";

                    string status = rs.Fields.Item("U_Status").Value.ToString();
                    string statusClass = status.ToLower();

                    opGroups[opName].Append($@"
                <div class='row'>
                    <div class='cell chassis'>{rs.Fields.Item("U_ChassisNo").Value}</div>
                    <div class='cell product' title='{rs.Fields.Item("U_FGDesc").Value}'>{rs.Fields.Item("U_FGDesc").Value}</div>
                    <div class='cell status'><span class='badge {statusClass}'>{status}</span></div>
                    <div class='cell date'>{startDate}</div>
                </div>");
                    rs.MoveNext();
                }

                var sortedKeys = new System.Collections.Generic.List<string>(opGroups.Keys);
                sortedKeys.Sort((x, y) => {
                    int GetPriority(string name)
                    {
                        string n = name.ToUpper();
                        if (n.Contains("TRIM")) return 1;
                        if (n.Contains("CHASIS")) return 2;
                        if (n.Contains("FINAL")) return 3;
                        if (n.Contains("TEST")) return 4;
                        return 5;
                    }
                    return GetPriority(x).CompareTo(GetPriority(y));
                });

                StringBuilder boxesHtml = new StringBuilder();
                foreach (string key in sortedKeys)
                {
                    boxesHtml.Append($@"
                <div class='op-card'>
                    <div class='op-header'>{key}</div>
                    <div class='grid-header'>
                        <div class='col col-chassis'>CHASSIS</div>
                        <div class='col col-product'>PRODUCT</div>
                        <div class='col col-status'>STATUS</div>
                        <div class='col col-date'>DATE</div>
                    </div>
                    <div class='scroll-box'>{opGroups[key].ToString()}</div>
                </div>");
                }

                // --- SECTION 2: WAITING (PENDING) ---
                string queryPending = @"
            SELECT X.""U_ChasisNo"", X.""U_RouteDsp"", X.""U_ProductName""
            FROM (
                SELECT T1.""U_ChasisNo"", T1.""LineId"", T1.""U_RouteDsp"", T2.""U_ProductName"",
                    ROW_NUMBER() OVER (PARTITION BY T1.""U_ChasisNo"" ORDER BY T1.""LineId"") AS ""RN""
                FROM ""@WRKORDRDTLSC"" T1
                INNER JOIN ""@WRKORDRDTLSH"" T2 ON T1.""DocEntry"" = T2.""DocEntry""
                WHERE T1.""U_Status"" = 'Pending'
                  AND EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S1 WHERE S1.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S1.""U_Status"" = 'Closed')
                  AND NOT EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S2 WHERE S2.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S2.""U_Status"" = 'WIP')
                  AND T1.""LineId"" > (SELECT MAX(""LineId"") FROM ""@WRKORDRDTLSC"" S3 WHERE S3.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S3.""U_Status"" = 'Closed')
            ) X WHERE X.""RN"" = 1 ORDER BY X.""U_ChasisNo""";

                rs.DoQuery(queryPending);
                StringBuilder pendingRows = new StringBuilder();
                while (!rs.EoF)
                {
                    pendingRows.Append($@"
                <div class='row'>
                    <div class='cell chassis'>{rs.Fields.Item("U_ChasisNo").Value}</div>
                    <div class='cell product' title='{rs.Fields.Item("U_ProductName").Value}'>{rs.Fields.Item("U_ProductName").Value}</div>
                    <div class='cell status' style='color: var(--accent); font-size:11px;'>{rs.Fields.Item("U_RouteDsp").Value}</div>
                </div>");
                    rs.MoveNext();
                }

                if (pendingRows.Length > 0)
                {
                    boxesHtml.Append($@"
                <div class='op-card' style='border: 1px dashed var(--accent);'>
                    <div class='op-header' style='color: #fff;'>WAITING AREA</div>
                    <div class='grid-header'>
                        <div class='col col-chassis'>CHASSIS</div>
                        <div class='col col-product'>PRODUCT</div>
                        <div class='col col-status'>NEXT STEP</div>
                    </div>
                    <div class='scroll-box'>{pendingRows.ToString()}</div>
                </div>");
                }

                // --- SECTION 3: GENERATE HTML (With Auto-Refresh Tag) ---
                string html = $@"<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta http-equiv='refresh' content='5'> <!-- REFRESH EVERY 5 SECONDS -->
    <title>LMC Dashboard</title>
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&display=swap');
        :root {{
            --bg: #0b0e14;
            --card-bg: rgba(30, 41, 59, 0.7);
            --accent: #38bdf8;
            --success: #22c55e;
            --warning: #eab308;
            --text-main: #f8fafc;
            --text-dim: #94a3b8;
        }}
        body {{ 
            margin: 0; padding: 40px; font-family: 'Inter', sans-serif; 
            background: radial-gradient(circle at top right, #1e293b, #0b0e14);
            color: var(--text-main); min-height: 100vh;
        }}
        h1 {{ text-align: center; font-weight: 800; font-size: 42px; margin-bottom: 50px; text-transform: uppercase; letter-spacing: -1px; background: linear-gradient(to bottom, #fff, #94a3b8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; }}
        .dashboard-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(650px, 1fr)); gap: 30px; }}
        .op-card {{ background: var(--card-bg); backdrop-filter: blur(12px); border-radius: 20px; border: 1px solid rgba(255,255,255,0.1); box-shadow: 0 25px 50px -12px rgba(0,0,0,0.5); display: flex; flex-direction: column; }}
        .op-header {{ padding: 20px 25px; font-size: 22px; font-weight: 800; color: var(--accent); border-bottom: 1px solid rgba(255,255,255,0.05); }}
        .grid-header {{ display: flex; padding: 12px 25px; background: rgba(0,0,0,0.2); font-size: 11px; font-weight: 700; color: var(--text-dim); letter-spacing: 1px; }}
        .scroll-box {{ max-height: 450px; overflow-y: auto; scrollbar-width: none; padding: 10px 0; }}
        .scroll-box::-webkit-scrollbar {{ display: none; }}
        .row {{ display: flex; align-items: center; padding: 15px 25px; border-bottom: 1px solid rgba(255,255,255,0.03); }}
        .cell {{ padding: 0 5px; overflow: hidden; }}
        .chassis {{ width: 30%; font-family: monospace; font-weight: 600; color: var(--accent); }}
        .product {{ width: 40%; font-size: 13px; color: #cbd5e1; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; line-height: 1.4; }}
        .status {{ width: 30%; text-align: center; }}
        .date {{ width: 15%; text-align: right; color: var(--text-dim); font-size: 12px; }}
        .badge {{ padding: 4px 10px; border-radius: 8px; font-size: 10px; font-weight: 800; display: inline-block; width: 65px; text-align: center; }}
        .badge.start {{ background: rgba(34,197,94,0.15); color: var(--success); border: 1px solid var(--success); }}
        .badge.pause {{ background: rgba(234,179,8,0.15); color: var(--warning); border: 1px solid var(--warning); }}
        .badge.resume {{ background: rgba(56,189,248,0.15); color: var(--accent); border: 1px solid var(--accent); }}
    </style>
</head>
<body>
    <h1>LMC Production Dashboard <span style='font-size: 12px; color: var(--text-dim);'>Updated: {DateTime.Now.ToString("HH:mm:ss")}</span></h1>
    <div class='dashboard-grid'>
        {boxesHtml}
    </div>
</body>
</html>";

                string filePath = Path.Combine(Path.GetTempPath(), "LMC_Dashboard.html");
                File.WriteAllText(filePath, html);

                // ONLY open the browser once. After that, just update the file content.
                if (!_isDashboardOpen)
                {
                    Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    _isDashboardOpen = true;

                    // Start the SAP Timer to refresh data every 5 seconds
                    StartDashboardTimer();
                }
            }
            catch (Exception ex) { SBO_Application.MessageBox(ex.Message); }
            finally { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }
        private void OpenTrimAssemblyStatusBoard(bool forceOpen = false)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                if (_Company == null || !_Company.Connected) return;
                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Define our 5 main sections to ensure they exist even if empty
                var sections = new System.Collections.Generic.Dictionary<string, StringBuilder> {
            { "TRIM", new StringBuilder() },
            { "CHASSIS", new StringBuilder() },
            { "FINAL", new StringBuilder() },
            { "TEST", new StringBuilder() },
            { "WAITING", new StringBuilder() }
        };

                var counts = new System.Collections.Generic.Dictionary<string, int> {
            { "TRIM", 0 }, { "CHASSIS", 0 }, { "FINAL", 0 }, { "TEST", 0 }, { "WAITING", 0 }
        };

                // --- SECTION 1: ACTIVE JOBS ---
                string queryActive = @"
                    SELECT
                    T0.""U_ChassisNo"", 
                    T0.""U_FGDesc"", 
                    T0.""U_JobDesc"", 
                    T0.""U_Status"", 
                    T0.""U_StartTime"",
                    CASE
                        WHEN T0.""U_Status"" = 'Start' THEN T1.""U_FirstBarcode""
                        WHEN T0.""U_Status"" = 'Resume' THEN T1.""U_JobResume""
                        WHEN T0.""U_Status"" = 'Pause' THEN T1.""U_JobPause""
                        ELSE T1.""U_LastBarcode""
                    END AS ""Current_Barcode""
                FROM ""@WJOBEXEH"" T0
                LEFT JOIN ""@WRKORDRDTLSC"" T1 ON T1.""U_WorkId"" = T0.""U_JobID""
                WHERE T0.""U_Status"" <> 'Stop'
                ORDER BY T0.""U_StartTime"" DESC";

                rs.DoQuery(queryActive);
                while (!rs.EoF)
                {
                    string opName = rs.Fields.Item("U_JobDesc").Value.ToString().ToUpper();
                    string currentBarcode = rs.Fields.Item("Current_Barcode").Value.ToString();
                    string targetKey = "TRIM"; // Default
                    if (opName.Contains("CHASIS") || opName.Contains("CHASSIS")) targetKey = "CHASSIS";
                    else if (opName.Contains("FINAL")) targetKey = "FINAL";
                    else if (opName.Contains("TEST")) targetKey = "TEST";
                    else targetKey = "TRIM";

                    string status = rs.Fields.Item("U_Status").Value.ToString().ToUpper();
                    // Map SAP Status to CSS Classes
                    string badgeClass = "badge-running";
                    if (status == "PAUSE" || status == "STOP") badgeClass = "badge-onhold";
                    if (status == "COMPLETE" || status == "CLOSED") badgeClass = "badge-complete";

                    string time = rs.Fields.Item("U_StartTime").Value.ToString();
                    if (time.Length == 4) time = time.Insert(2, ":");

                    counts[targetKey]++;
                    sections[targetKey].Append($@"
                <div class='t-row prod-col-tpl'>
                    <span class='t-sn'>{counts[targetKey]}</span>
                    <span class='t-vin'>{rs.Fields.Item("U_ChassisNo").Value}</span>
                    <span class='t-model' title='{rs.Fields.Item("U_FGDesc").Value}'>{rs.Fields.Item("U_FGDesc").Value}</span>
                    <span class='t-badge {badgeClass}'>{status}</span>
                    <span class='t-station stn-{targetKey.ToLower()}'>{currentBarcode}</span>
                    <span class='t-time'>{time}</span>
                    <button class='t-del-btn'>×</button>
                </div>");
                    rs.MoveNext();
                }

                // --- SECTION 2: WAITING (PENDING) ---
                string queryPending = @"
            SELECT X.""U_ChasisNo"", X.""U_RouteDsp"", X.""U_ProductName""
            FROM (
                SELECT T1.""U_ChasisNo"", T1.""LineId"", T1.""U_RouteDsp"", T2.""U_ProductName"",
                    ROW_NUMBER() OVER (PARTITION BY T1.""U_ChasisNo"" ORDER BY T1.""LineId"") AS ""RN""
                FROM ""@WRKORDRDTLSC"" T1
                INNER JOIN ""@WRKORDRDTLSH"" T2 ON T1.""DocEntry"" = T2.""DocEntry""
                WHERE T1.""U_Status"" = 'Pending'
                  AND EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S1 WHERE S1.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S1.""U_Status"" = 'Closed')
                  AND NOT EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S2 WHERE S2.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S2.""U_Status"" = 'WIP')
                  AND T1.""LineId"" > (SELECT MAX(""LineId"") FROM ""@WRKORDRDTLSC"" S3 WHERE S3.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S3.""U_Status"" = 'Closed')
            ) X WHERE X.""RN"" = 1 ORDER BY X.""U_ChasisNo""";

                rs.DoQuery(queryPending);
                while (!rs.EoF)
                {
                    counts["WAITING"]++;
                    sections["WAITING"].Append($@"
                <div class='t-row wait-col-tpl'>
                    <span class='t-sn'>{counts["WAITING"]}</span>
                    <span class='t-vin'>{rs.Fields.Item("U_ChasisNo").Value}</span>
                    <span class='t-model' title='{rs.Fields.Item("U_ProductName").Value}'>{rs.Fields.Item("U_ProductName").Value}</span>
                    <span class='t-badge badge-waiting'>WAITING</span>
                    <div class='t-last-station'><div class='ls-dot'></div><span class='ls-val'>{rs.Fields.Item("U_RouteDsp").Value}</span></div>
                    <button class='t-del-btn'>×</button>
                </div>");
                    rs.MoveNext();
                }

                int totalInProd = counts["TRIM"] + counts["CHASSIS"] + counts["FINAL"] + counts["TEST"];
                int totalVins = totalInProd + counts["WAITING"];

                // Helper to handle empty states
                foreach (var key in sections.Keys.ToList())
                {
                    if (sections[key].Length == 0)
                        sections[key].Append($"<div class='t-empty'>No VINs in {key}</div>");
                }

                // --- GENERATE MODERN HTML ---
                string html = $@"
<!DOCTYPE html>
<html lang='en'>
<head>
    <meta charset='UTF-8'>
    <meta http-equiv='refresh' content='10'>
    <title>LMC Production Dashboard</title>
    <link href='https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800&family=Barlow:wght@400;500;600;700&family=Chakra+Petch:wght@500;600;700&display=swap' rel='stylesheet'>
    <style>
        /* Embed the CSS from the provided design */
        *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Barlow', sans-serif; background: #f0f2f5; color: #0f172a; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }}
        .topbar {{ background: #002c5f; border-bottom: 3px solid #00aad2; height: 56px; display: flex; align-items: center; justify-content: space-between; padding: 0 28px; flex-shrink: 0; }}
        .topbar-plant .plant-name {{ font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 15px; letter-spacing: 0.08em; text-transform: uppercase; color: #fff; }}
        .topbar-plant .plant-sub {{ font-size: 10px; color: #7dd3fc; text-transform: uppercase; }}
        .stat-pill {{ text-align: right; padding: 0 18px; border-left: 1px solid rgba(255,255,255,0.12); }}
        .sp-val {{ font-family: 'Chakra Petch', monospace; font-size: 22px; font-weight: 700; line-height: 1; }}
        .v-cyan {{ color: #00aad2; }} .v-green {{ color: #4ade80; }} .v-red {{ color: #f87171; }}
        .sp-label {{ font-family: 'Chakra Petch', monospace; font-size: 8px; letter-spacing: 0.2em; text-transform: uppercase; color: #7dd3fc; }}
        .flowbar {{ background: #fff; border-bottom: 1px solid #e2e8f0; padding: 0 28px; height: 34px; display: flex; align-items: center; flex-shrink: 0; gap: 15px; }}
        .flow-name {{ font-family: 'Barlow Condensed', sans-serif; font-size: 12px; font-weight: 700; text-transform: uppercase; }}
        .flow-count {{ font-family: 'Chakra Petch', monospace; font-size: 9px; padding: 1px 7px; border-radius: 10px; background: #f1f5f9; }}
        .body-wrap {{ flex: 1; display: flex; flex-direction: column; gap: 10px; padding: 12px 16px; overflow: hidden; }}
        .line-grid {{ flex: 1; display: grid; grid-template-columns: 1fr 1fr; gap: 10px; min-height: 0; }}
        .section-panel {{ background: #fff; border-radius: 4px; box-shadow: 0 4px 16px rgba(0,0,0,0.05); display: flex; flex-direction: column; overflow: hidden; border-top: 4px solid #002c5f; }}
        .panel-head {{ background: #0f172a; padding: 8px 16px; display: flex; align-items: center; justify-content: space-between; }}
        .ph-name {{ font-family: 'Barlow Condensed', sans-serif; font-weight: 700; color: #fff; text-transform: uppercase; font-size: 14px; }}
        .ph-count {{ font-family: 'Chakra Petch', monospace; font-size: 9px; padding: 2px 9px; border-radius: 20px; background: #dbeafe; color: #002c5f; }}
        .col-head {{ background: #f8fafc; border-bottom: 1px solid #e2e8f0; padding: 5px 16px; display: grid; gap: 8px; }}
        .col-head span {{ font-family: 'Chakra Petch', monospace; font-size: 8px; color: #94a3b8; text-transform: uppercase; }}
        .prod-col-tpl {{ grid-template-columns: 26px 140px minmax(0,1fr) 90px 58px 58px 20px; }}
        .wait-col-tpl {{ grid-template-columns: 26px 140px minmax(0,1fr) 90px minmax(0,1fr) 20px; }}
        .table-rows {{ flex: 1; overflow-y: auto; }}
        .t-row {{ padding: 6px 16px; border-bottom: 1px solid #f1f5f9; display: grid; align-items: center; gap: 8px; }}
        .t-vin {{ font-family: 'Chakra Petch', monospace; font-size: 10px; font-weight: 600; color: #002c5f; }}
        .t-model {{ font-size: 11px; font-weight: 500; color: #334155; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .t-badge {{ font-family: 'Barlow Condensed', sans-serif; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 3px; text-transform: uppercase; text-align:center; }}
        .badge-running {{ background: #dcfce7; color: #15803d; }}
        .badge-onhold {{ background: #fef3c7; color: #b45309; }}
        .badge-waiting {{ background: #fee2e2; color: #dc2626; }}
        .ls-val {{ font-family: 'Chakra Petch', monospace; font-size: 12px; font-weight: 700; color: #dc2626; }}
        .t-empty {{ text-align: center; padding: 20px; font-size: 10px; color: #cbd5e1; font-family: 'Chakra Petch'; }}
        .stn-trim {{ color: #002c5f; }} .stn-chassis {{ color: #d97706; }} .stn-final {{ color: #16a34a; }} .stn-test {{ color: #7c3aed; }}
        .stn-trim, .stn-chassis, .stn-final, .stn-test {{font - family: 'Chakra Petch', monospace; 
            font-size: 10px; 
            font-weight: 700;
            white-space: nowrap; /* Prevents long barcodes from breaking into two lines */
        }}
        .page-footer {{ background: #fff; border-top: 1px solid #e2e8f0; padding: 0 28px; height: 30px; display: flex; align-items: center; justify-content: space-between; font-size: 8px; color: #94a3b8; font-family: 'Chakra Petch'; text-transform: uppercase; }}
    </style>
</head>
<body>
    <div class='topbar'>
        <div class='topbar-plant'>
            <div class='plant-name'>Laxmi Motor Corporation</div>
            <div class='plant-sub'>VIN Production Tracking &nbsp;·&nbsp; Ramgram Plant</div>
        </div>
        <div style='display:flex;'>
            <div class='stat-pill'><div class='sp-val v-cyan'>{totalVins:00}</div><div class='sp-label'>Total VINs</div></div>
            <div class='stat-pill'><div class='sp-val v-green'>{totalInProd:00}</div><div class='sp-label'>In Production</div></div>
            <div class='stat-pill'><div class='sp-val v-red'>{counts["WAITING"]:00}</div><div class='sp-label'>Waiting</div></div>
        </div>
    </div>

    <div class='flowbar'>
        <span class='flow-name'>Trim <span class='flow-count'>{counts["TRIM"]}</span></span> ›
        <span class='flow-name'>Chassis <span class='flow-count'>{counts["CHASSIS"]}</span></span> ›
        <span class='flow-name'>Final <span class='flow-count'>{counts["FINAL"]}</span></span> ›
        <span class='flow-name'>Test <span class='flow-count'>{counts["TEST"]}</span></span> ›
        <span class='flow-name' style='color:#dc2626'>Waiting <span class='flow-count'>{counts["WAITING"]}</span></span>
    </div>

    <div class='body-wrap'>
        <div class='line-grid'>
            <div class='section-panel'>
                <div class='panel-head'><span class='ph-name'>Trim Line</span><span class='ph-count'>{counts["TRIM"]}</span></div>
                <div class='col-head prod-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Stn</span><span>Time</span><span></span></div>
                <div class='table-rows'>{sections["TRIM"]}</div>
            </div>
            <div class='section-panel' style='border-top-color:#d97706'>
                <div class='panel-head'><span class='ph-name'>Chassis Line</span><span class='ph-count'>{counts["CHASSIS"]}</span></div>
                <div class='col-head prod-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Stn</span><span>Time</span><span></span></div>
                <div class='table-rows'>{sections["CHASSIS"]}</div>
            </div>
        </div>
        <div class='line-grid'>
            <div class='section-panel' style='border-top-color:#16a34a'>
                <div class='panel-head'><span class='ph-name'>Final Line</span><span class='ph-count'>{counts["FINAL"]}</span></div>
                <div class='col-head prod-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Stn</span><span>Time</span><span></span></div>
                <div class='table-rows'>{sections["FINAL"]}</div>
            </div>
            <div class='section-panel' style='border-top-color:#7c3aed'>
                <div class='panel-head'><span class='ph-name'>Test Track</span><span class='ph-count'>{counts["TEST"]}</span></div>
                <div class='col-head prod-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Stn</span><span>Time</span><span></span></div>
                <div class='table-rows'>{sections["TEST"]}</div>
            </div>
        </div>
        <div class='section-panel' style='border-top-color:#dc2626; flex: 0 0 180px;'>
            <div class='panel-head'><span class='ph-name'>Waiting Area</span><span class='ph-count'>{counts["WAITING"]}</span></div>
            <div class='col-head wait-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Next Required Station</span><span></span></div>
            <div class='table-rows'>{sections["WAITING"]}</div>
        </div>
    </div>

    <div class='page-footer'>
        <span>LMC · Hyundai Authorized Assembler · Ramgram, Nepal</span>
        <span>Updated: {DateTime.Now:HH:mm:ss}</span>
    </div>
</body>
</html>";

                string filePath = Path.Combine(Path.GetTempPath(), "LMC_Dashboard.html");
                File.WriteAllText(filePath, html);

                if (!_isDashboardOpen || forceOpen)
                {
                    //Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
                    //System.Diagnostics.Process.Start("msedge.exe", filePath);
                    System.Diagnostics.Process.Start("msedge.exe", "\"" + filePath + "\"");
                    _isDashboardOpen = true;
                    StartDashboardTimer();
                }
            }
            catch (Exception ex) { SBO_Application.MessageBox(ex.Message); }
            finally { if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs); }
        }


        private void OpenTrimAssemblyStatusBoard2(bool forceOpen = false)
        {
            SAPbobsCOM.Recordset rs = null;
            SAPbobsCOM.Recordset rsUser = null;
            try
            {
                if (_Company == null || !_Company.Connected) return;

                // 1. Get the current User's assigned station and Normalize spelling
                string userStation = "ALL";
                try
                {
                    rsUser = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    string userQuery = $@"SELECT ""U_Station"" FROM ""OUSR"" WHERE ""USERID"" = '{_Company.UserSignature}'";
                    rsUser.DoQuery(userQuery);
                    if (!rsUser.EoF)
                    {
                        userStation = rsUser.Fields.Item("U_Station").Value.ToString().Trim().ToUpper();

                        // NORMALIZE: If DB has "CHASIS", treat it as "CHASSIS" to match the rest of the logic
                        if (userStation == "CHASIS") userStation = "CHASSIS";
                    }
                }
                catch { userStation = "ALL"; }

                rs = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Initialize Section Buffers
                var sections = new System.Collections.Generic.Dictionary<string, StringBuilder> {
            { "TRIM", new StringBuilder() },
            { "CHASSIS", new StringBuilder() },
            { "FINAL", new StringBuilder() },
            { "TEST", new StringBuilder() },
            { "WAITING", new StringBuilder() }
        };

                var counts = new System.Collections.Generic.Dictionary<string, int> {
            { "TRIM", 0 }, { "CHASSIS", 0 }, { "FINAL", 0 }, { "TEST", 0 }, { "WAITING", 0 }
        };

                // 2. Fetch Active Jobs
                string queryActive = @"
            SELECT T0.""U_ChassisNo"", T0.""U_FGDesc"", T0.""U_JobDesc"", T0.""U_Status"", T0.""U_StartTime"",
            CASE
                WHEN T0.""U_Status"" = 'Start' THEN T1.""U_FirstBarcode""
                WHEN T0.""U_Status"" = 'Resume' THEN T1.""U_JobResume""
                WHEN T0.""U_Status"" = 'Pause' THEN T1.""U_JobPause""
                ELSE T1.""U_LastBarcode""
            END AS ""Current_Barcode""
            FROM ""@WJOBEXEH"" T0
            LEFT JOIN ""@WRKORDRDTLSC"" T1 ON T1.""U_WorkId"" = T0.""U_JobID""
            WHERE T0.""U_Status"" <> 'Stop'
            ORDER BY T0.""U_StartTime"" DESC";

                rs.DoQuery(queryActive);
                while (!rs.EoF)
                {
                    string opName = rs.Fields.Item("U_JobDesc").Value.ToString().ToUpper();
                    string currentBarcode = rs.Fields.Item("Current_Barcode").Value.ToString();

                    // Map the DB Operation to one of our 4 sections
                    string targetKey = "TRIM"; // Default
                    if (opName.Contains("CHASIS") || opName.Contains("CHASSIS")) targetKey = "CHASSIS";
                    else if (opName.Contains("FINAL")) targetKey = "FINAL";
                    else if (opName.Contains("TEST")) targetKey = "TEST";

                    string status = rs.Fields.Item("U_Status").Value.ToString().ToUpper();
                    string badgeClass = (status == "PAUSE") ? "badge-onhold" : "badge-running";

                    string time = rs.Fields.Item("U_StartTime").Value.ToString();
                    if (time.Length == 4) time = time.Insert(2, ":");

                    counts[targetKey]++;
                    sections[targetKey].Append($@"
                <div class='t-row prod-col-tpl'>
                    <span class='t-sn'>{counts[targetKey]}</span>
                    <span class='t-vin'>{rs.Fields.Item("U_ChassisNo").Value}</span>
                    <span class='t-model' title='{rs.Fields.Item("U_FGDesc").Value}'>{rs.Fields.Item("U_FGDesc").Value}</span>
                    <span class='t-badge {badgeClass}'>{status}</span>
                    <span class='t-station stn-{targetKey.ToLower()}'>{currentBarcode}</span>
                    <span class='t-time'>{time}</span>
                    <button class='t-del-btn'>×</button>
                </div>");
                    rs.MoveNext();
                }

                // 3. Fetch Waiting Area Jobs
                string queryPending = @"
            SELECT X.""U_ChasisNo"", X.""U_RouteDsp"", X.""U_ProductName""
            FROM (
                SELECT T1.""U_ChasisNo"", T1.""LineId"", T1.""U_RouteDsp"", T2.""U_ProductName"",
                    ROW_NUMBER() OVER (PARTITION BY T1.""U_ChasisNo"" ORDER BY T1.""LineId"") AS ""RN""
                FROM ""@WRKORDRDTLSC"" T1
                INNER JOIN ""@WRKORDRDTLSH"" T2 ON T1.""DocEntry"" = T2.""DocEntry""
                WHERE T1.""U_Status"" = 'Pending'
                  AND EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S1 WHERE S1.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S1.""U_Status"" = 'Closed')
                  AND NOT EXISTS (SELECT 1 FROM ""@WRKORDRDTLSC"" S2 WHERE S2.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S2.""U_Status"" = 'WIP')
                  AND T1.""LineId"" > (SELECT MAX(""LineId"") FROM ""@WRKORDRDTLSC"" S3 WHERE S3.""U_ChasisNo"" = T1.""U_ChasisNo"" AND S3.""U_Status"" = 'Closed')
            ) X WHERE X.""RN"" = 1 ORDER BY X.""U_ChasisNo""";

                rs.DoQuery(queryPending);
                while (!rs.EoF)
                {
                    counts["WAITING"]++;
                    sections["WAITING"].Append($@"
                <div class='t-row wait-col-tpl'>
                    <span class='t-sn'>{counts["WAITING"]}</span>
                    <span class='t-vin'>{rs.Fields.Item("U_ChasisNo").Value}</span>
                    <span class='t-model' title='{rs.Fields.Item("U_ProductName").Value}'>{rs.Fields.Item("U_ProductName").Value}</span>
                    <span class='t-badge badge-waiting'>WAITING</span>
                    <div class='t-last-station'><div class='ls-dot'></div><span class='ls-val'>{rs.Fields.Item("U_RouteDsp").Value}</span></div>
                    <button class='t-del-btn'>×</button>
                </div>");
                    rs.MoveNext();
                }

                int totalInProd = counts["TRIM"] + counts["CHASSIS"] + counts["FINAL"] + counts["TEST"];
                int totalVins = totalInProd + counts["WAITING"];

                // 4. Build Dynamic Content HTML
                StringBuilder dynamicContent = new StringBuilder();
                string gridStyle = (userStation == "ALL") ? "grid-template-columns: 1fr 1fr;" : "grid-template-columns: 1fr;";

                // Layout Helper Method
                string BuildSection(string title, string key, string color)
                {
                    string content = (sections[key].Length > 0) ? sections[key].ToString() : $"<div class='t-empty'>No Active VINs in {title}</div>";
                    return $@"<div class='section-panel' style='border-top-color:{color}'>
                        <div class='panel-head'><span class='ph-name'>{title}</span><span class='ph-count'>{counts[key]}</span></div>
                        <div class='col-head prod-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Stn</span><span>Time</span><span></span></div>
                        <div class='table-rows'>{content}</div>
                      </div>";
                }

                // Row 1: TRIM and CHASSIS
                if (userStation == "ALL" || userStation == "TRIM" || userStation == "CHASSIS")
                {
                    dynamicContent.Append($@"<div class='line-grid' style='{gridStyle}'>");
                    if (userStation == "ALL" || userStation == "TRIM") dynamicContent.Append(BuildSection("Trim Line", "TRIM", "#002c5f"));
                    if (userStation == "ALL" || userStation == "CHASSIS") dynamicContent.Append(BuildSection("Chassis Line", "CHASSIS", "#d97706"));
                    dynamicContent.Append("</div>");
                }

                // Row 2: FINAL and TEST
                if (userStation == "ALL" || userStation == "FINAL" || userStation == "TEST")
                {
                    string marginTop = (userStation == "ALL") ? "margin-top:10px;" : "";
                    dynamicContent.Append($@"<div class='line-grid' style='{marginTop} {gridStyle}'>");
                    if (userStation == "ALL" || userStation == "FINAL") dynamicContent.Append(BuildSection("Final Line", "FINAL", "#16a34a"));
                    if (userStation == "ALL" || userStation == "TEST") dynamicContent.Append(BuildSection("Test Track", "TEST", "#7c3aed"));
                    dynamicContent.Append("</div>");
                }

                // Always show Waiting Area
                dynamicContent.Append($@"
            <div class='section-panel' style='border-top-color:#dc2626; margin-top:10px; flex: 0 0 180px;'>
                <div class='panel-head'><span class='ph-name'>Waiting Area</span><span class='ph-count'>{counts["WAITING"]}</span></div>
                <div class='col-head wait-col-tpl'><span>#</span><span>VIN</span><span>Model</span><span>Status</span><span>Next Required Station</span><span></span></div>
                <div class='table-rows'>{(counts["WAITING"] > 0 ? sections["WAITING"].ToString() : "<div class='t-empty'>No VINs Waiting</div>")}</div>
            </div>");

                // 5. Generate Full HTML Output
                string html = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='UTF-8'>
    <meta http-equiv='refresh' content='10'>
    <title>LMC Production Dashboard</title>
    <link href='https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@600;700;800&family=Barlow:wght@400;500;600;700&family=Chakra+Petch:wght@500;600;700&display=swap' rel='stylesheet'>
    <style>
        *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
        body {{ font-family: 'Barlow', sans-serif; background: #f0f2f5; color: #0f172a; display: flex; flex-direction: column; height: 100vh; overflow: hidden; }}
        .topbar {{ background: #002c5f; border-bottom: 3px solid #00aad2; height: 56px; display: flex; align-items: center; justify-content: space-between; padding: 0 28px; flex-shrink: 0; }}
        .topbar-plant .plant-name {{ font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 15px; letter-spacing: 0.08em; text-transform: uppercase; color: #fff; }}
        .topbar-plant .plant-sub {{ font-size: 10px; color: #7dd3fc; text-transform: uppercase; }}
        .stat-pill {{ text-align: right; padding: 0 18px; border-left: 1px solid rgba(255,255,255,0.12); }}
        .sp-val {{ font-family: 'Chakra Petch', monospace; font-size: 22px; font-weight: 700; line-height: 1; }}
        .v-cyan {{ color: #00aad2; }} .v-green {{ color: #4ade80; }} .v-red {{ color: #f87171; }}
        .sp-label {{ font-family: 'Chakra Petch', monospace; font-size: 8px; letter-spacing: 0.2em; text-transform: uppercase; color: #7dd3fc; }}
        .body-wrap {{ flex: 1; display: flex; flex-direction: column; gap: 10px; padding: 12px 16px; overflow: hidden; }}
        .line-grid {{ display: grid; gap: 10px; min-height: 0; flex: 1; }}
        .section-panel {{ background: #fff; border-radius: 4px; box-shadow: 0 4px 16px rgba(0,0,0,0.05); display: flex; flex-direction: column; overflow: hidden; border-top: 4px solid #002c5f; }}
        .panel-head {{ background: #0f172a; padding: 8px 16px; display: flex; align-items: center; justify-content: space-between; }}
        .ph-name {{ font-family: 'Barlow Condensed', sans-serif; font-weight: 700; color: #fff; text-transform: uppercase; font-size: 14px; }}
        .ph-count {{ font-family: 'Chakra Petch', monospace; font-size: 9px; padding: 2px 9px; border-radius: 20px; background: #dbeafe; color: #002c5f; }}
        .col-head {{ background: #f8fafc; border-bottom: 1px solid #e2e8f0; padding: 5px 16px; display: grid; gap: 8px; }}
        .col-head span {{ font-family: 'Chakra Petch', monospace; font-size: 8px; color: #94a3b8; text-transform: uppercase; }}
        .prod-col-tpl {{ grid-template-columns: 26px 140px minmax(0,1fr) 90px 80px 58px 20px; }}
        .wait-col-tpl {{ grid-template-columns: 26px 140px minmax(0,1fr) 90px minmax(0,1fr) 20px; }}
        .table-rows {{ flex: 1; overflow-y: auto; }}
        .t-row {{ padding: 6px 16px; border-bottom: 1px solid #f1f5f9; display: grid; align-items: center; gap: 8px; }}
        .t-vin {{ font-family: 'Chakra Petch', monospace; font-size: 10px; font-weight: 600; color: #002c5f; }}
        .t-model {{ font-size: 11px; font-weight: 500; color: #334155; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
        .t-badge {{ font-family: 'Barlow Condensed', sans-serif; font-size: 10px; font-weight: 700; padding: 2px 6px; border-radius: 3px; text-transform: uppercase; text-align:center; }}
        .badge-running {{ background: #dcfce7; color: #15803d; }}
        .badge-onhold {{ background: #fef3c7; color: #b45309; }}
        .badge-waiting {{ background: #fee2e2; color: #dc2626; }}
        .ls-val {{ font-family: 'Chakra Petch', monospace; font-size: 12px; font-weight: 700; color: #dc2626; }}
        .t-empty {{ text-align: center; padding: 20px; font-size: 11px; color: #94a3b8; font-family: 'Chakra Petch'; }}
        .stn-trim {{ color: #002c5f; }} .stn-chassis {{ color: #d97706; }} .stn-final {{ color: #16a34a; }} .stn-test {{ color: #7c3aed; }}
        .t-station {{ font-family: 'Chakra Petch', monospace; font-size: 10px; font-weight: 700; }}
        .page-footer {{ background: #fff; border-top: 1px solid #e2e8f0; padding: 0 28px; height: 30px; display: flex; align-items: center; justify-content: space-between; font-size: 8px; color: #94a3b8; font-family: 'Chakra Petch'; text-transform: uppercase; }}
    </style>
</head>
<body>
    <div class='topbar'>
        <div class='topbar-plant'>
            <div class='plant-name'>Laxmi Motor Corporation</div>
            <div class='plant-sub'>STATION: {userStation} &nbsp;·&nbsp; Ramgram Plant</div>
        </div>
        <div style='display:flex;'>
            <div class='stat-pill'><div class='sp-val v-cyan'>{totalVins:00}</div><div class='sp-label'>Total VINs</div></div>
            <div class='stat-pill'><div class='sp-val v-green'>{totalInProd:00}</div><div class='sp-label'>In Production</div></div>
            <div class='stat-pill'><div class='sp-val v-red'>{counts["WAITING"]:00}</div><div class='sp-label'>Waiting</div></div>
        </div>
    </div>
    <div class='body-wrap'>
        {dynamicContent.ToString()}
    </div>
    <div class='page-footer'>
        <span>LMC · Ramgram, Nepal</span>
        <span>Last Update: {DateTime.Now:HH:mm:ss}</span>
    </div>
</body>
</html>";

                string filePath = Path.Combine(Path.GetTempPath(), "LMC_Dashboard.html");
                File.WriteAllText(filePath, html);

                if (!_isDashboardOpen || forceOpen)
                {
                    System.Diagnostics.Process.Start("msedge.exe", "\"" + filePath + "\"");
                    _isDashboardOpen = true;
                    StartDashboardTimer();
                }
            }
            catch (Exception ex) { SBO_Application.MessageBox(ex.Message); }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                if (rsUser != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsUser);
            }
        }

        // Timer setup to call the function automatically
        private void StartDashboardTimer()
        {
            if (_refreshTimer == null)
            {
                _refreshTimer = new System.Windows.Forms.Timer();
                _refreshTimer.Interval = 5000; // 5 seconds
                _refreshTimer.Tick += (s, e) => {
                    OpenTrimAssemblyStatusBoard();
                };
                _refreshTimer.Start();
            }
        }
        private void StartFocusTimer(SAPbouiCOM.Form form)
        {
            _oForm = form; // Assign form reference
            focusTimer = new System.Windows.Forms.Timer();
            focusTimer.Interval = 200; // 0.2 sec
            focusTimer.Tick += FocusTimer_Tick;
            focusTimer.Start();
        }
        private void FocusTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (_oForm != null && _oForm.Visible)
                {
                    _oForm.Items.Item("txtScan").Specific.Focus();
                }
                else
                {
                    // Stop timer if form is closed
                    focusTimer.Stop();
                }
            }
            catch
            {
                focusTimer.Stop();
            }
        }


        private void SetRemarkFieldsState(SAPbouiCOM.Form oForm, string status)
        {
            oForm.Freeze(true);
            try
            {
                if (status == "Start")
                {
                    oForm.Items.Item("txtRmkIn").Enabled = true;   // Editable
                    oForm.Items.Item("lblRmkOut").Visible = false; // Hidden
                    oForm.Items.Item("txtRmkOut").Visible = false; // Hidden
                }
                else if (status == "Pause")
                {
                    oForm.Items.Item("txtRmkIn").Enabled = false;  // Read only
                    oForm.Items.Item("lblRmkOut").Visible = false; // Hidden
                    oForm.Items.Item("txtRmkOut").Visible = false; // Hidden
                }
                else if (status == "Stop")
                {
                    oForm.Items.Item("txtRmkIn").Enabled = false;  // Read only
                    oForm.Items.Item("lblRmkOut").Visible = true;  // Show
                    oForm.Items.Item("txtRmkOut").Visible = true;  // Show
                    oForm.Items.Item("txtRmkOut").Enabled = true;  // Editable
                }
            }
            catch { }
            finally { oForm.Freeze(false); }
        }

        private void LoadFinishedJobUI(string docEntry)
        {
            try
            {
                // 1. Create/Load the Job Execution Form
                this._Object = new ClsJobOrderExecution();
                Utilities.LoadForm(ref this._Object, Constants.Forms.JobOrderExecutions);

                SAPbouiCOM.Form oJobForm = ((ClsJobOrderExecution)this._Object).Form;

                if (oJobForm != null)
                {
                    oJobForm.Freeze(true);
                    try
                    {
                        // 2. Put form in Find Mode
                        oJobForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                        // 3. Set DocEntry and Click Find
                        // By now, HANA has updated the status to 'Stop'
                        ((SAPbouiCOM.EditText)oJobForm.Items.Item("txtDocE").Specific).Value = docEntry;
                        oJobForm.Items.Item("1").Click();

                        // 4. Apply UI Logic (Show Out Remarks, Lock In Remarks)
                        // This method will now see the 'Stop' status from the database
                        SetRemarkFieldsState(oJobForm, "Stop");
                    }
                    catch (Exception exInner)
                    {
                        _Application.StatusBar.SetText("Find Record Error: " + exInner.Message);
                    }
                    finally
                    {
                        oJobForm.Freeze(false);
                    }
                }
            }
            catch (Exception ex)
            {
                _Application.StatusBar.SetText("Form Load Error: " + ex.Message);
            }
        }


        #region Login for Service Layer
        private async Task<string> LoginToServiceLayer()
        {
            // Reuse session if available
            if (!string.IsNullOrEmpty(_slCookieHeader)) return _slCookieHeader;

            System.Net.ServicePointManager.Expect100Continue = false;
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (m, c, ch, e) => true,
                AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate,
                UseCookies = false
            };

            using (var client = new HttpClient(handler))
            {
                // Exact JSON from your Postman test
                string rawJson = "{\"CompanyDB\":\"TEST_MESI\",\"UserName\":\"SUPPORT02\",\"Password\":\"1@34%\"}";
                var utf8WithoutBom = new System.Text.UTF8Encoding(false);
                var content = new ByteArrayContent(utf8WithoutBom.GetBytes(rawJson));
                content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

                client.DefaultRequestHeaders.Add("Accept", "*/*");
                client.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.32.3");

                var response = await client.PostAsync(_slUrl + "Login", content);

                if (response.IsSuccessStatusCode)
                {
                    // Capture B1SESSION and ROUTEID cookies
                    if (response.Headers.TryGetValues("Set-Cookie", out var cookies))
                    {
                        _slCookieHeader = string.Join("; ", cookies);
                    }
                    return _slCookieHeader;
                }
                else
                {
                    string error = await response.Content.ReadAsStringAsync();
                    throw new Exception("Login Failed: " + error);
                }
            }
        }
        #endregion

        #region Create issue for production via service layer
        public async Task<int> Create_Issue_ServiceLayer(string prodEntry, string routeId, string workId, string chassisNo, string jobExeNo)
        {
            try
            {
                // 1. Ensure we are logged in and have cookies
                string cookie = await LoginToServiceLayer();

                // 2. Fetch BOM and Mapping data via DI-API Recordset (Fast Read-Only)
                var lineItems = new List<object>();

                string bomQuery = $@"
            SELECT T0.""ItemCode"", T0.""PlannedQty"", T0.""LineNum"", T1.""PlannedQty"" as ""headerQty"", T2.""ManSerNum"", T2.""U_SerialType""
            FROM WOR1 T0 
            INNER JOIN OWOR T1 ON T0.""DocEntry"" = T1.""DocEntry""
            INNER JOIN OITM T2 on T0.""ItemCode"" = T2.""ItemCode""
            WHERE T0.""DocEntry"" = '{prodEntry}' AND T0.""StageId"" = '{routeId}' AND T2.""InvntItem"" = 'Y'";

                SAPbobsCOM.Recordset rsBOM = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsBOM.DoQuery(bomQuery);

                SAPbobsCOM.Recordset rsMap = (SAPbobsCOM.Recordset)_Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsMap.DoQuery($@"SELECT * FROM ""@ENGCHASISMMC"" WHERE ""U_ChasisNo"" = '{chassisNo}'");

                if (rsMap.EoF) throw new Exception($"Chassis mapping not found for {chassisNo}");

                while (!rsBOM.EoF)
                {
                    double hQty = Convert.ToDouble(rsBOM.Fields.Item("headerQty").Value);
                    double pQty = Convert.ToDouble(rsBOM.Fields.Item("PlannedQty").Value);
                    double qtyToIssue = (hQty > 0) ? (pQty / hQty) : 0;

                    var serialsList = new List<object>();

                    // Handle Serial Managed Items
                    if (rsBOM.Fields.Item("ManSerNum").Value.ToString() == "Y")
                    {
                        string sType = rsBOM.Fields.Item("U_SerialType").Value.ToString();
                        string mappedSerial = "";

                        if (sType == "ENGINE") mappedSerial = rsMap.Fields.Item("U_EngineNo").Value.ToString();
                        else if (sType == "CHASSIS") mappedSerial = rsMap.Fields.Item("U_ChasisNo").Value.ToString();
                        else if (sType == "TRANSMISSION") mappedSerial = rsMap.Fields.Item("U_TransNo").Value.ToString();
                        else if (sType == "KEY") mappedSerial = rsMap.Fields.Item("U_SetKey").Value.ToString();

                        if (!string.IsNullOrEmpty(mappedSerial))
                        {
                            serialsList.Add(new { InternalSerialNumber = mappedSerial, Quantity = 1.0 });
                            qtyToIssue = 1.0;
                        }
                    }

                    lineItems.Add(new
                    {
                        BaseEntry = Convert.ToInt32(prodEntry),
                        BaseType = 202,
                        BaseLine = Convert.ToInt32(rsBOM.Fields.Item("LineNum").Value),
                        Quantity = qtyToIssue,
                        SerialNumbers = serialsList
                    });

                    rsBOM.MoveNext();
                }

                // 3. Prepare Final Payload
                var issuePayload = new
                {
                    DocDate = DateTime.Now.ToString("yyyy-MM-dd"),
                    Comments = $"Auto Issue (ServiceLayer) | Chassis: {chassisNo} | WorkID: {workId}",
                    U_ProdEntry = prodEntry,
                    U_JobExeNo = jobExeNo,
                    DocumentLines = lineItems
                };

                // 4. Send POST Request
                var handler = new HttpClientHandler
                {
                    ServerCertificateCustomValidationCallback = (m, c, ch, e) => true,
                    AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate,
                    UseCookies = false
                };

                using (var client = new HttpClient(handler))
                {
                    client.Timeout = TimeSpan.FromMinutes(10);

                    client.DefaultRequestHeaders.Add("Cookie", cookie);
                    client.DefaultRequestHeaders.Add("Accept", "*/*");
                    client.DefaultRequestHeaders.Add("User-Agent", "PostmanRuntime/7.32.3");

                    string jsonBody = JsonConvert.SerializeObject(issuePayload);
                    var utf8WithoutBom = new System.Text.UTF8Encoding(false);
                    var content = new ByteArrayContent(utf8WithoutBom.GetBytes(jsonBody));
                    content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/json");

                    var response = await client.PostAsync(_slUrl + "InventoryGenExits", content);

                    if (response.IsSuccessStatusCode)
                    {
                        var responseContent = await response.Content.ReadAsStringAsync();
                        dynamic result = JsonConvert.DeserializeObject(responseContent);
                        return (int)result.DocEntry;
                    }
                    else
                    {
                        string errorBody = await response.Content.ReadAsStringAsync();

                        // If session expired (301), clear cookie for next attempt
                        if (errorBody.Contains("301")) _slCookieHeader = "";

                        throw new Exception(errorBody);
                    }
                }
            }
            catch (Exception ex)
            {
                _Application.StatusBar.SetText("Service Layer Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return 0;
            }
        }
        #endregion


    }
}
