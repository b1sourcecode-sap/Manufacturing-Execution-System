
using Sap.Data.Hana;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;


namespace Production_Planning_LMC
{
    class ProductionPlanningLMC : Base
    {

        #region VARIABLE DECLARATION
        SAPbouiCOM.Item _ExistItem, _NewItem, _Items;
        private bool flag;
        public SAPbouiCOM.Matrix _Matrix;
        SAPbouiCOM.StaticText _StaticText;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "";
        string _PRNum = string.Empty, _PRName = string.Empty, _RevNum = string.Empty;
        public string ErrMsg = string.Empty;
        public int ErrCode = 0;
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        private static HanaConnection _conn;
        private static HanaDataAdapter _da;
        //private static DataTable _dataTable;
        private static HanaParameter _parm1;
        private static HanaParameter _parm2;
        private static int _errors = 0;
        string strcon, _Sql;
        private SAPbouiCOM.EditText _EditText;
        private SAPbouiCOM.LinkedButton _Linked;
        private SAPbouiCOM.ComboBox _ComboBoxM;
        private SAPbouiCOM.ComboBox _ComboBoxC;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        private SAPbouiCOM.ComboBox _ComboBoxH;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null, _MasterDBDataSource1 = null;
        private SAPbouiCOM.ChooseFromListCollection _CFLS;
        private SAPbouiCOM.ChooseFromList _CFL;
        private SAPbouiCOM.ChooseFromListCreationParams _CFLC;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        private SAPbouiCOM.Form _UDFForm;
        private string _ActivityType = string.Empty;
        string _ItemCode = string.Empty, _ItemName = string.Empty;
        private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSPP, rsUpdate;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        SAPbouiCOM.DBDataSource _DBData = null;
        private System.Collections.Generic.List<string> ListOfBlockedID;
        string selectedLotNo;
        string itmGrpCode, model, itmGrpCode1;
        string toCreateQty;
        private string _deferLotNo = "";
        SAPbobsCOM.Recordset _RS1 = null;
        private bool _deferFind = false; // Flag to indicate a find operation is pending
        private bool _isHandlingCFLEvent = false;
        private int _rowToDelete = -1;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public ProductionPlanningLMC()
            : base()
        {

        }
        ~ProductionPlanningLMC()
        {
        }
        #endregion

        #region FormDataEvent
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            base.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);

            // This event runs AFTER a form loads existing data.
            if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
            {
                this.Form.Freeze(true);
                Utilities.SerializedMartix(ref _Form, ref _Matrix);
                UpdatePushToInventoryDisplay();
                EnsureRowsForRequiredQty();
                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    Disable();
                }
                this.Form.Freeze(false);
            }

        }
        #endregion

        #region Menu Event
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            base.Menu_Event(ref pVal, ref BubbleEvent);

            #region Menu_Event BeforeAction
            if (pVal.BeforeAction)
            {
                switch (pVal.MenuUID)
                {
                    case "1293":
                        DeleteRowAndResequence();
                        BubbleEvent = false;
                        break;

                    case Constants.System_Menus.mnu_FIRST:

                        break;
                    case Constants.System_Menus.mnu_LAST:

                        break;
                    case Constants.System_Menus.mnu_PREVIOUS:

                        break;
                    case Constants.System_Menus.mnu_NEXT:

                        break;


                }
            }
            #endregion

            #region Menu_Event After Action
            else
            {
                //if (pVal.MenuUID == "1293")
                //{
                //    //UpdateMatrixAndQuantitiesAfterDelete();
                //    DeleteRowAndResequence();

                //}
            }
            #endregion
        }
        #endregion

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {
                    #region Item Pressed
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && (_Form.Mode == BoFormMode.fm_ADD_MODE || _Form.Mode == BoFormMode.fm_UPDATE_MODE))
                        {
                            if (!Validation())
                            {
                                BubbleEvent = false;
                            }

                            try
                            {
                                _Form.Freeze(true);

                                // Ensure the matrix data is synchronized with the data source.
                                _Matrix.FlushToDataSource();

                                int createdQty = 0;

                                // Iterate through the child data source to count valid rows.
                                // The last row might be an empty "new entry" row, so we check for content.
                                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                                {
                                    string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", i).Trim();
                                    string modelCode = _ChildDBDataSource.GetValue("U_ModelCode", i).Trim();

                                    //update form, when pincode is not yet gebdafe
                                    //string pinCode = _ChildDBDataSource.GetValue("U_PinCode", i).Trim();
                                    //if (!string.IsNullOrEmpty(engineNo))
                                    //{ 
                                    //    if (_Form.Mode == BoFormMode.fm_UPDATE_MODE)
                                    //    {
                                    //        if (string.IsNullOrEmpty(pinCode))
                                    //        {
                                    //            Utilities.Application.SBO_Application.SetStatusBarMessage($"Row {i + 1}: Pin Code is required before updating.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    //            BubbleEvent = false;
                                    //            return;
                                    //        }
                                    //    }
                                    //}

                                        // A row is considered valid if key fields are not empty.
                                        // The Validation() method already ensures that if engineNo exists, the rest of the required fields also exist.
                                        if (!string.IsNullOrEmpty(engineNo) && !string.IsNullOrEmpty(modelCode))
                                    {
                                        createdQty++;
                                    }
                                }
                                double toCreateQty = 0;
                                string toCreateQtyStr = _MasterDBDataSource.GetValue("U_ToCreate", 0).Trim();
                                if (!string.IsNullOrEmpty(toCreateQtyStr) && double.TryParse(toCreateQtyStr, out double parsedQty))
                                {
                                    toCreateQty = parsedQty;
                                }

                                // Calculate the remaining quantity.
                                double remainingQty = toCreateQty - createdQty;

                                _MasterDBDataSource.SetValue("U_CreatedQty", 0, createdQty.ToString());
                                _MasterDBDataSource.SetValue("U_RemQty", 0, remainingQty.ToString());
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error during quantity calculation: " + ex.Message);
                                BubbleEvent = false; // Prevent the Add/Update operation if there's an error.
                            }
                            finally
                            {
                                _Form.Freeze(false);
                            }
                        }

                        break;
                    #endregion

                    #region CFL
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "txtLotNo")
                        {
                            this.FilterItemType("");
                        }
                        if (pVal.ColUID == "colEngine")
                        {
                            this.FilterEngineType(pVal.Row);
                        }
                        if (pVal.ColUID == "colChasis")
                        {
                            this.FilterChasisType(pVal.Row);
                        }
                        if (pVal.ColUID == "colTrans")
                        {
                            this.FilterTransType(pVal.Row);
                        }
                        if (pVal.ColUID == "colSetKey")
                        {
                            this.FilterKeyType(pVal.Row);
                        }
                        if (pVal.ColUID == "colModel")
                        {
                            this.FilterModel(pVal.Row);
                        }

                        break;
                    #endregion

                    #region Push to Inventory Click
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "prodMtrx" && pVal.ColUID == "colInvt" && pVal.Row > 0)
                        {
                            SAPbouiCOM.ProgressBar progBar = null;

                            try
                            {
                                string cellValue = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colInvt").Cells.Item(pVal.Row).Specific).Value.Trim();
                                if (cellValue == "Push to Inventory")
                                {
                                    int _DocEntry = Convert.ToInt32(((SAPbouiCOM.EditText)_Matrix.Columns.Item("colProdNo").Cells.Item(pVal.Row).Specific).Value.Trim());
                                    string _Chaisis = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colChasis").Cells.Item(pVal.Row).Specific).Value.Trim();
                                    string _Engine = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colEngine").Cells.Item(pVal.Row).Specific).Value.Trim();
                                    string _Key = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colSetKey").Cells.Item(pVal.Row).Specific).Value.Trim();
                                    string _Transimission = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colTrans").Cells.Item(pVal.Row).Specific).Value.Trim();

                                    int confirm = Utilities.Application.SBO_Application.MessageBox($"Create Receipt from Production for Chassis {_Chaisis}?", 1, "Yes", "No");

                                    if (confirm == 1)
                                    {
                                        //Create_ReceiptFrom_Production(_DocEntry, _Chaisis,_Engine,_Key,_Transimission);
                                        progBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Receipt from Production........", 1, false);

                                        int receiptDocEntry = Create_ReceiptFrom_Production(_DocEntry, _Chaisis, _Engine, _Key, _Transimission);

                                        if (receiptDocEntry > 0)
                                        {
                                            UpdateRowAfterReceipt(_DocEntry.ToString(), _Chaisis, receiptDocEntry);
                                            _Matrix.FlushToDataSource();

                                            int dsIndex = pVal.Row - 1;

                                            _ChildDBDataSource.SetValue("U_PushInvt", dsIndex, "");

                                            _ChildDBDataSource.SetValue("U_RepProd", dsIndex, receiptDocEntry.ToString());
                                            //if receipt from production is updated (update the status to Manufactured)
                                            //string _ReceiptFrmProd = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colRecpP").Cells.Item(pVal.Row).Specific).Value.Trim();
                                            string _ReceiptFrmProd = _ChildDBDataSource.GetValue("U_RepProd", dsIndex);
                                            if (!string.IsNullOrEmpty(_ReceiptFrmProd))
                                            {
                                                _ChildDBDataSource.SetValue("U_Status", dsIndex, "Manufactured");
                                            }

                                            // _ChildDBDataSource.SetValue("U_PushStatus", dsIndex, "Completed"); 

                                            _Matrix.LoadFromDataSource();

                                            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                            {
                                                _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                            }
                                            _Form.Items.Item("1").Click();
                                            //refresh form after updating 
                                            _Form.Refresh();
                                        }
                                        //progBar.Stop();
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error in Push to Inventory: " + ex.Message);
                            }

                            finally
                            {
                                // 3. The Finally block ensures the UI is unfrozen even if an error occurs
                                if (progBar != null)
                                {
                                    try
                                    {
                                        progBar.Stop();
                                    }
                                    catch { /* Ignore errors during stop */ }

                                    // Best Practice: Release COM object to prevent memory leaks in SAP
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(progBar);
                                    progBar = null;
                                }
                            }
                        }

                        break;
                        #endregion
                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    #region Set CFL Values

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                        // LotNo CFL
                        if (pVal.ItemUID == "txtLotNo" && _SysCFLEvent.SelectedObjects != null)
                        {
                            if (_isHandlingCFLEvent)
                            {
                                return; // Exit if we are already in this handler
                            }

                            _isHandlingCFLEvent = true; // Set the flag
                            try
                            {
                                _Form.Freeze(true);
                                //selectedLotNo = _SysCFLEvent.SelectedObjects.GetValue("U_LotNo", 0).ToString().Trim();
                                //_MasterDBDataSource.SetValue("U_LotNo", 0, selectedLotNo);
                                selectedLotNo = _SysCFLEvent.SelectedObjects.GetValue("Code", 0).ToString().Trim();

                                // Check how many records exist
                                SAPbobsCOM.Recordset rsLot = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string queryLot = $@"SELECT COUNT(*) AS ""CNT"", MAX(""DocEntry"") AS ""DocEntry"" FROM ""@ENGCHASISMMH"" WHERE ""U_LotNo"" = '{selectedLotNo.Replace("'", "''")}'";
                                rsLot.DoQuery(queryLot);
                                int count = Convert.ToInt32(rsLot.Fields.Item("CNT").Value);

                                SAPbouiCOM.EditText lotNoField = (SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific;

                                System.Windows.Forms.Application.DoEvents();
                                System.Threading.Thread.Sleep(50);

                                if (count == 0)
                                {
                                    _MasterDBDataSource.SetValue("U_LotNo", 0, selectedLotNo);

                                    // Fetch GRPO

                                    string query = $@"SELECT ""DocEntry"", ""U_Model"" FROM OPDN WHERE ""U_LotNo"" = '{selectedLotNo}'";
                                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    rs.DoQuery(query);


                                    //string queryQty = $@"SELECT count(""DistNumber"") FROM OSRN  WHERE ""LotNumber"" = '{selectedLotNo}' and ""U_SerialType"" = 'CHASSIS'";
                                    //SAPbobsCOM.Recordset rs2 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    //rs2.DoQuery(queryQty);
                                    if (!rs.EoF)
                                    {
                                        //string docEntry = rs.Fields.Item("DocEntry").Value.ToString();
                                        //model = rs.Fields.Item("U_Model").Value.ToString();
                                        //_MasterDBDataSource.SetValue("U_GRPO", 0, docEntry);
                                        //_MasterDBDataSource.SetValue("U_Model", 0, model);
                                        var grpoDocEntries = new System.Collections.Generic.List<string>();

                                        //each lot number have same model code
                                        model = rs.Fields.Item("U_Model").Value.ToString();

                                        // Loop through all records returned by the query.
                                        while (!rs.EoF)
                                        {
                                            grpoDocEntries.Add(rs.Fields.Item("DocEntry").Value.ToString());
                                            rs.MoveNext();
                                        }

                                        // Join the list into a single, comma-separated string.
                                        string concatenatedGrpos = string.Join(",", grpoDocEntries);

                                        _MasterDBDataSource.SetValue("U_GRPO", 0, concatenatedGrpos);
                                        _MasterDBDataSource.SetValue("U_Model", 0, model);
                                        _MasterDBDataSource.SetValue("U_BatchNo", 0, "");


                                        string query1 = $@"SELECT ""ItmsGrpCod"" FROM OITB WHERE ""ItmsGrpNam"" = '{model}'";
                                        SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rs1.DoQuery(query1);
                                        if (!rs1.EoF)
                                        {
                                            itmGrpCode = rs1.Fields.Item("ItmsGrpCod").Value.ToString();
                                        }
                                        else
                                        {
                                            Utilities.ShowErrorMessage("No Item Group Code found for Model: " + model);
                                        }

                                        string queryQty = $@"SELECT count(""DistNumber"") AS ""Qty"" FROM OSRN  WHERE ""LotNumber"" = '{selectedLotNo}' and ""U_SerialType"" = 'CHASSIS'";
                                        SAPbobsCOM.Recordset rs2 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rs2.DoQuery(queryQty);
                                        if (!rs2.EoF)
                                        {
                                            toCreateQty = rs2.Fields.Item("Qty").Value.ToString();
                                            _MasterDBDataSource.SetValue("U_ToCreate", 0, toCreateQty);
                                            _MasterDBDataSource.SetValue("U_NoChasis", 0, toCreateQty);
                                        }
                                        else
                                        {
                                            //Utilities.ShowErrorMessage("No Item Group Code found for Model: " + model);
                                        }
                                    }
                                    else
                                    {
                                        Utilities.ShowErrorMessage("No GRPO found for Lot No: " + selectedLotNo);
                                    }


                                    // Prepare Matrix for user input
                                    _Matrix.Clear();
                                    _ChildDBDataSource.Clear();
                                    _ChildDBDataSource.InsertRecord(0);

                                    _Matrix.LoadFromDataSource();
                                    AddSequenceNumbersToMatrix(_Matrix);

                                    EnsureRowsForRequiredQty();
                                }
                                else if (count == 1)
                                {
                                    //string codeQuery = $@"SELECT ""DocEntry"" FROM ""@ENGCHASISMMH"" WHERE ""U_LotNo"" = '{selectedLotNo.Replace("'", "''")}'";
                                    //rsLot.DoQuery(codeQuery);
                                    //string recordCode = rsLot.Fields.Item("DocEntry").Value.ToString();

                                    //SAPbobsCOM.GeneralService oService = Utilities.Application.Company.GetCompanyService().GetGeneralService("ENGCHASISMM");
                                    //SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                                    //oParams.SetProperty("DocEntry", recordCode);
                                    //SAPbobsCOM.GeneralData oData = oService.GetByParams(oParams);

                                    //_MasterDBDataSource.SetValue("U_LotNo", 0, oData.GetProperty("U_LotNo").ToString());
                                    //_MasterDBDataSource.SetValue("U_GRPONo", 0, oData.GetProperty("U_GRPONo").ToString());
                                    //_MasterDBDataSource.SetValue("U_DocNo", 0, oData.GetProperty("U_DocNo").ToString());
                                    //_MasterDBDataSource.SetValue("U_Status", 0, oData.GetProperty("U_Status").ToString());
                                    //_MasterDBDataSource.SetValue("U_PostDate", 0, oData.GetProperty("U_PostDate").ToString("yyyyMMdd"));
                                    //_MasterDBDataSource.SetValue("U_ReqQty", 0, oData.GetProperty("U_ReqQty").ToString());
                                    //_MasterDBDataSource.SetValue("U_CreatedQty", 0, oData.GetProperty("U_CreatedQty").ToString());
                                    //_MasterDBDataSource.SetValue("U_RemQty", 0, oData.GetProperty("U_RemQty").ToString());
                                    //_MasterDBDataSource.SetValue("U_Model", 0, oData.GetProperty("U_Model").ToString());
                                    //_MasterDBDataSource.SetValue("U_ToCreate", 0, oData.GetProperty("U_ToCreate").ToString());

                                    ////Populate matrix
                                    //string queryChild = $@"Select ""U_SrNo"", ""U_EngineNo"", ""U_ChasisNo"", ""U_TransNo"", ""U_SetKey"", ""U_ModelCode"", 
                                    // ""U_Status"", ""U_PushInvt"", ""U_RepProd"", ""U_ProdOrdNo"" 
                                    // from ""@ENGCHASISMMC"" where ""DocEntry"" = '{recordCode}' and ""U_EngineNo"" IS NOT NULL ORDER BY CAST(""U_SrNo"" AS INTEGER) ASC";
                                    //Utilities.ExecuteSQL(ref _RS1, queryChild);

                                    //for (int i = 0; i < _RS1.RecordCount; i++)
                                    //{
                                    //    //_Matrix.FlushToDataSource();
                                    //    _ChildDBDataSource.InsertRecord(i);
                                    //    _ChildDBDataSource.SetValue("U_SrNo", i, _RS1.Fields.Item("U_SrNo").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_EngineNo", i, _RS1.Fields.Item("U_EngineNo").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_ChasisNo", i, _RS1.Fields.Item("U_ChasisNo").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_TransNo", i, _RS1.Fields.Item("U_TransNo").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_SetKey", i, _RS1.Fields.Item("U_SetKey").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_ModelCode", i, _RS1.Fields.Item("U_ModelCode").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_Status", i, _RS1.Fields.Item("U_Status").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_PushInvt", i, _RS1.Fields.Item("U_PushInvt").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_RepProd", i, _RS1.Fields.Item("U_RepProd").Value.ToString());
                                    //    _ChildDBDataSource.SetValue("U_ProdOrdNo", i, _RS1.Fields.Item("U_ProdOrdNo").Value.ToString());
                                    //    _Matrix.LoadFromDataSource();
                                    //    _RS1.MoveNext();
                                    //}


                                    //// _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    ////Change the form Mode
                                    //_Form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;



                                    SAPbobsCOM.Recordset rsLot1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string codeQuery1 = $@"SELECT ""DocEntry"", ""U_LotNo"" 
                                     FROM ""@ENGCHASISMMH"" 
                                     WHERE ""U_LotNo"" = '{selectedLotNo.Replace("'", "''")}'";

                                    rsLot1.DoQuery(codeQuery1);
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                    string recordLotNo = rsLot1.Fields.Item("U_LotNo").Value.ToString();

                                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific).Value = recordLotNo;
                                    _MasterDBDataSource.SetValue("U_LotNo", 0, recordLotNo);

                                    System.Threading.Thread.Sleep(60);

                                    _Form.Items.Item("1").Click();

                                    System.Threading.Thread.Sleep(60);

                                    // find method reload the data 
                                    string osrnQty = $@"SELECT COUNT(""DistNumber"") AS ""Qty"" FROM OSRN WHERE ""LotNumber"" = '{selectedLotNo}' and ""U_SerialType"" = 'CHASSIS'";
                                    SAPbobsCOM.Recordset rsQtyCheck = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    rsQtyCheck.DoQuery(osrnQty);

                                    if (!rsQtyCheck.EoF)
                                    {
                                        double latestOsrnQty = Convert.ToInt32(rsQtyCheck.Fields.Item("Qty").Value);

                                        int actualCreatedQty = 0;
                                        for (int i = 0; i < _ChildDBDataSource.Size; i++)
                                        {
                                            string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", i).Trim();
                                            if (!string.IsNullOrEmpty(engineNo))
                                            {
                                                actualCreatedQty++;
                                            }
                                        }

                                        //check additional grpo numbers
                                        SAPbobsCOM.Recordset rs = null;
                                        string query = $@"SELECT ""DocEntry"" FROM OPDN WHERE ""U_LotNo"" = '{recordLotNo.Replace("'", "''")}'";
                                        rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rs.DoQuery(query);

                                        var grpoDocEntries = new System.Collections.Generic.List<string>();

                                        while (!rs.EoF)
                                        {
                                            grpoDocEntries.Add(rs.Fields.Item("DocEntry").Value.ToString());
                                            rs.MoveNext();
                                        }

                                        string concatenatedGrpos = string.Join(",", grpoDocEntries);

                                        // Get the current value on screen to compare
                                        string currentOnScreen = _MasterDBDataSource.GetValue("U_GRPO", 0).Trim();


                                        //int docToCreateQty = Convert.ToInt32(_MasterDBDataSource.GetValue("U_ToCreate", 0).Trim());
                                        //int docCreatedQty = Convert.ToInt32(_MasterDBDataSource.GetValue("U_CreatedQty", 0).Trim());

                                        double docToCreateQty = Convert.ToDouble(_MasterDBDataSource.GetValue("U_ToCreate", 0));
                                        double docNoChasis = Convert.ToDouble(_MasterDBDataSource.GetValue("U_NoChasis", 0));
                                        double docCreatedQty = Convert.ToDouble(_MasterDBDataSource.GetValue("U_CreatedQty", 0));

                                        if (latestOsrnQty != docToCreateQty || actualCreatedQty != docCreatedQty || latestOsrnQty != docNoChasis || currentOnScreen != concatenatedGrpos)
                                        {
                                            _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                                            _MasterDBDataSource.SetValue("U_ToCreate", 0, latestOsrnQty.ToString());
                                            _MasterDBDataSource.SetValue("U_NoChasis", 0, latestOsrnQty.ToString());
                                            _MasterDBDataSource.SetValue("U_CreatedQty", 0, actualCreatedQty.ToString());
                                            _MasterDBDataSource.SetValue("U_GRPO", 0, concatenatedGrpos);

                                            double newRemainingQty = latestOsrnQty - actualCreatedQty;
                                            _MasterDBDataSource.SetValue("U_RemQty", 0, newRemainingQty.ToString());

                                            EnsureRowsForRequiredQty();

                                            _Form.Items.Item("1").Click();
                                        }
                                        //UpdateGRPOForLot(recordLotNo);
                                    }

                                }
                                else
                                {
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                                    lotNoField.Value = selectedLotNo;

                                    //small delay
                                    System.Windows.Forms.Application.DoEvents();
                                    System.Threading.Thread.Sleep(50);

                                    _Form.Items.Item("1").Click();
                                }
                                //if (_Form.Mode  == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                //{
                                //    _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                //}
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error while setting LotNo CFL value: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);

                                System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
                                {
                                    try
                                    {
                                        SAPbouiCOM.Cell firstCell = _Matrix.Columns.Item("colChasis").Cells.Item(1);
                                        firstCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        ((SAPbouiCOM.EditText)firstCell.Specific).Active = true;
                                    }
                                    catch { }
                                });
                                _isHandlingCFLEvent = false;

                            }
                        }

                        // CFL for matrix columns
                        if (pVal.ItemUID == "prodMtrx" && _SysCFLEvent.SelectedObjects != null)
                        {
                            _Form.Freeze(true);
                            try
                            {
                                //string value = _SysCFLEvent.SelectedObjects.GetValue("DistNumber", 0).ToString().Trim();
                                //string itm = _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim();
                          

                                //string modelNameQuery = $@"SELECT ""ItemName"" from OITM
                                //WHERE ""ItemCode"" = '{rsModel.Fields.Item("U_ItemCode").Value}'";
                                //SAPbobsCOM.Recordset rsModelName = null;
                                //rsModelName = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //rsModelName.DoQuery(modelNameQuery);


                                _Matrix.FlushToDataSource();

                                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item(_SysCFLEvent.ChooseFromListUID);
                                string objectType = oCFL.ObjectType;


                                string valueField = objectType == "4" ? "ItemCode" : "DistNumber";

                                string value = SafeGetValue(_SysCFLEvent.SelectedObjects, valueField);
                                string itm = SafeGetValue(_SysCFLEvent.SelectedObjects, "ItemCode");
                                string itmName = SafeGetValue(_SysCFLEvent.SelectedObjects, "ItemName");

                                string currentLotNo = _MasterDBDataSource.GetValue("U_LotNo", 0).Trim();
                                string modelQuery = $@"SELECT distinct T0.""U_ItemCode"", T1.""ItemName"" from OSRN T0 
                                INNER JOIN OITM T1 ON T0.""U_ItemCode"" = T1.""ItemCode""
                                WHERE T0.""U_SerialType"" = 'CHASSIS' AND T0.""LotNumber"" = '{currentLotNo}' and T0.""DistNumber"" = '{value}' ";
                                SAPbobsCOM.Recordset rsModel = null;
                                rsModel = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                rsModel.DoQuery(modelQuery);

                                int msRow = pVal.Row - 1;

                                switch (pVal.ColUID)
                                {
                                    case "colEngine":
                                        //_ChildDBDataSource.SetValue("U_EngineNo", _ChildDBDataSource.Offset, value);
                                        //_ChildDBDataSource.SetValue("U_Status", _ChildDBDataSource.Offset, "Available");
                                        _ChildDBDataSource.SetValue("U_EngineNo", msRow, value);
                                        _ChildDBDataSource.SetValue("U_Status", msRow, "Available");
                                        _ChildDBDataSource.SetValue("U_PushInvt", msRow, "");
                                        _Matrix.SetLineData(pVal.Row);

                                        int currentRow = pVal.Row;
                                        MoveFocusToNext("colSetKey", currentRow);
                                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                        }

                                        break;

                                    case "colChasis":
                                        _ChildDBDataSource.SetValue("U_ChasisNo", msRow, value);
                                        _ChildDBDataSource.SetValue("U_ModelCode", msRow, rsModel.Fields.Item("U_ItemCode").Value.ToString());
                                        _ChildDBDataSource.SetValue("U_ModelName", msRow, rsModel.Fields.Item("ItemName").Value.ToString());
                                        _Matrix.SetLineData(pVal.Row);

                                        int currentRowC = pVal.Row;
                                        MoveFocusToNext("colEngine", currentRowC);
                                        //System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
                                        //{
                                        //    try
                                        //    {
                                        //        SAPbouiCOM.Cell nextCell = _Matrix.Columns.Item("colEngine").Cells.Item(currentRowC);
                                        //        nextCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        //        ((SAPbouiCOM.EditText)nextCell.Specific).Active = true;
                                        //    }
                                        //    catch { }
                                        //});
                                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                        }
                                        break;

                                    case "colTrans":
                                        _ChildDBDataSource.SetValue("U_TransNo", msRow, value);
                                        int currentRowT = pVal.Row;
                                        MoveFocusToNext("colModel", currentRowT);
                                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                        }
                                        break;
                                    case "colSetKey":
                                        _ChildDBDataSource.SetValue("U_SetKey", msRow, value);
                                        _Matrix.SetLineData(pVal.Row);

                                        //AddNewRow();
                                        //int nextRow = pVal.Row + 1;
                                        int currentRowS = pVal.Row;
                                        MoveFocusToNext("colTrans", currentRowS);
                                        string reqQty = _MasterDBDataSource.GetValue("U_ToCreate", 0).Trim();
                                        double.TryParse(reqQty, out double ReqQ);
                                        if (pVal.Row < ReqQ)
                                        {
                                            if (currentRowS - 1 == _ChildDBDataSource.Size - 1)
                                            {
                                                // This is the true "new entry" row.
                                                // Now it's safe to add another blank row and move the cursor.
                                                AddNewRow();

                                                int nextRowToFocus = pVal.Row + 1;
                                                MoveFocusToNext("colModel", currentRowS);
                                            }
                                        }
                                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                        }

                                        break;
                                    case "colModel":
                                        int dsRow = pVal.Row - 1;

                                        _ChildDBDataSource.SetValue("U_ModelCode", dsRow, itm);
                                        _ChildDBDataSource.SetValue("U_ModelName", dsRow, itmName);
                                        _Matrix.SetLineData(pVal.Row);

                                        //string reqQty = _MasterDBDataSource.GetValue("U_ToCreate", 0).Trim();
                                        //double.TryParse(reqQty, out double ReqQ);
                                        //if (pVal.Row < ReqQ)
                                        //{
                                        //    if (dsRow == _ChildDBDataSource.Size - 1)
                                        //    {
                                        //        // This is the true "new entry" row.
                                        //        // Now it's safe to add another blank row and move the cursor.
                                        //        AddNewRow();

                                        //        int nextRowToFocus = pVal.Row + 1;
                                        //        MoveFocusToNext("colChasis", nextRowToFocus);
                                        //    }
                                        //}
                                        MoveFocusToNext("colModel", pVal.Row);
                                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                                        {
                                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                                        }
                                        break;

                                }

                                _Matrix.LoadFromDataSource();

                                // Only add new row after last CFL (SetKey)
                                //if (pVal.ColUID == "colSetKey")
                                //{
                                //    AddNewRow();
                                //}
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error while setting matrix CFL value: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);
                            }
                        }

                        break;


                    //case SAPbouiCOM.BoEventTypes.et_MENU_CLICK:
                    //    SAPbouiCOM.MenuEvent menuEvent = (SAPbouiCOM.MenuEvent)pVal;
                    //    if (menuEvent.BeforeAction == false)
                    //    {
                    //        // Check for "Delete Row" AND that the event came from your matrix
                    //        if (menuEvent.MenuUID == "1293" && pVal.ItemUID == "prodMtrx")
                    //        {
                    //            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("prodMtrx").Specific;
                    //            int clickedRow = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);

                    //            if (clickedRow > 0)
                    //            {
                    //                int dsRow = clickedRow - 1;

                    //                // Prevent crash if user deletes the blank "new entry" row
                    //                if (dsRow >= _ChildDBDataSource.Size) return;

                    //                string prodEnt = _ChildDBDataSource.GetValue("U_ProdOrdNo", dsRow).Trim();

                    //                if (string.IsNullOrEmpty(prodEnt))
                    //                {
                    //                    _Form.Freeze(true);
                    //                    try
                    //                    {
                    //                        // 1. Remove the record. The UI will update automatically.
                    //                        _ChildDBDataSource.RemoveRecord(dsRow);

                    //                        // 2. Update sequencing and quantities on the modified datasource.
                    //                        UpdateMatrixAndQuantitiesAfterDelete();

                    //                        if (_Form.Mode == BoFormMode.fm_OK_MODE)
                    //                        {
                    //                            _Form.Mode = BoFormMode.fm_UPDATE_MODE;
                    //                        }
                    //                    }
                    //                    catch (Exception ex) { Utilities.ShowErrorMessage("Error deleting row: " + ex.Message); }
                    //                    finally { _Form.Freeze(false); }
                    //                }
                    //                else
                    //                {
                    //                    Utilities.ShowErrorMessage("This row cannot be deleted; it is linked to a Production Order.");
                    //                }
                    //            }
                    //        }
                    //    }
                    //    break;


                    #endregion

                    #region  Item Pressed (Button Generate PinCode and Open Production Order)
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "btnGenPin" && (_Form.Mode == BoFormMode.fm_UPDATE_MODE || _Form.Mode == BoFormMode.fm_ADD_MODE))
                        {
                            GenerateVINPinFromMatrix();
                        }

                        if (pVal.ItemUID == "btnProdO" && (_Form.Mode == BoFormMode.fm_OK_MODE))
                        {

                            Utilities.Application.SBO_Application.ActivateMenuItem("4369");
                            _Form.Close();
                        }
                        break;
                        #endregion

                }
            }
            #endregion
        }
        #endregion

        #region FILTER Item Type
        internal void FilterItemType(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                DateTime parsedDate = DateTime.Now;
                string date = parsedDate.ToString("yyyyMMdd");
                //string endDt = _Form.DataSources.DBDataSources.Item("LOTNOMASTER").GetValue("U_EndDt", 0).Trim();
                _Cons = new Conditions();

                _Con = _Cons.Add();
                _Con.Alias = "U_Status";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Active";

                _Con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "U_EndDt";
                _Con.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_EQUAL;
                _Con.CondVal = date;

                _Con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;

                _Con = _Cons.Add();
                _Con.Alias = "U_EndDt";
                _Con.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL; // Checks if db value is NULL
                _Con.BracketCloseNum = 1;

                _Form.ChooseFromLists.Item("CFL_LotNo").SetConditions(_Cons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region FILTER Chassis Type
        internal void FilterChasisType(int currentRow)
        {

            _Form.Freeze(true);
            try
            {
                string currentLotNo = _MasterDBDataSource.GetValue("U_LotNo", 0).Trim();
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con;
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_ChasisNo");

                // Base Conditions
                con = oCons.Add();
                con.Alias = "LotNumber";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = currentLotNo;
                con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "U_SerialType";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "CHASSIS";
                con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "Quantity";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "1";
                con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "QuantOut";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "0";
                //con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                var chasisToExclude = new System.Collections.Generic.List<string>();
                for (int i = 1; i <= _Matrix.RowCount; i++)
                {
                    if (i == currentRow) continue;

                    string chasis = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colChasis").Cells.Item(i).Specific).Value.Trim();
                    if (!string.IsNullOrEmpty(chasis))
                    {
                        chasisToExclude.Add(chasis);
                    }
                }

                if (chasisToExclude.Count > 0)
                {
                    con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    for (int i = 0; i < chasisToExclude.Count; i++)
                    {
                        con = oCons.Add();
                        con.Alias = "DistNumber";
                        if (chasisToExclude[i].Contains("*"))
                        {
                            con.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_CONTAIN;
                        }
                        else
                        {
                            con.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        }
                        con.CondVal = chasisToExclude[i];

                        if (i < chasisToExclude.Count - 1)
                            con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    }
                }

                oCFL.SetConditions(oCons);

            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Filter Engine Type
        internal void FilterEngineType(int currentRow)
        {
            _Form.Freeze(true);
            try
            {
                string currentLotNo = _MasterDBDataSource.GetValue("U_LotNo", 0).Trim();
                //string lotNumber = ((SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific).Value.Trim();
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_EngineNo");
                //oCFL.ClearSelected(); // reset previous selection

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con;

                con = oCons.Add();
                con.Alias = "LotNumber";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = currentLotNo;
                con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "U_SerialType";
                con.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                con.CondVal = "ENGINE";
                con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "Quantity";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "1";
                con.Relationship = BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "QuantOut";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "0";


                // First, collect all engine numbers that need to be excluded
                var enginesToExclude = new System.Collections.Generic.List<string>();
                for (int i = 1; i <= _Matrix.RowCount; i++)
                {
                    if (i == currentRow) continue; // skip current row

                    string eng = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colEngine").Cells.Item(i).Specific).Value.Trim();
                    if (!string.IsNullOrEmpty(eng))
                    {
                        enginesToExclude.Add(eng);
                    }
                }

                // Now, add the exclusion conditions to the collection
                if (enginesToExclude.Count > 0)
                {
                    // Link the previous condition (U_SerialType) to this new block of conditions
                    con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    for (int i = 0; i < enginesToExclude.Count; i++)
                    {
                        con = oCons.Add();
                        con.Alias = "DistNumber";
                        con.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        con.CondVal = enginesToExclude[i];

                        // Set 'AND' relationship only if it is NOT the last item in the list
                        if (i < enginesToExclude.Count - 1)
                        {
                            con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        }
                    }
                }

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region FILTER Transmission Type
        internal void FilterTransType(int currentRow)
        {

            _Form.Freeze(true);
            try
            {
                string currentLotNo = _MasterDBDataSource.GetValue("U_LotNo", 0).Trim();
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_TransNo");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con;
                //_Cons = new Conditions();

                con = oCons.Add();
                //con.BracketOpenNum = 1;
                con.Alias = "LotNumber";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = currentLotNo;
                //con.BracketCloseNum = 1;
                con.Relationship = BoConditionRelationship.cr_AND;

                con = oCons.Add();
                //con.BracketOpenNum = 1;
                con.Alias = "U_SerialType";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "TRANSMISSION";
                con.Relationship = BoConditionRelationship.cr_AND;
                //con.BracketCloseNum = 1;

                con = oCons.Add();
                con.Alias = "Quantity";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "1";
                con.Relationship = BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "QuantOut";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "0";

                var enginesToExclude = new System.Collections.Generic.List<string>();
                for (int i = 1; i <= _Matrix.RowCount; i++)
                {
                    if (i == currentRow) continue; // skip current row

                    string eng = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colTrans").Cells.Item(i).Specific).Value.Trim();
                    if (!string.IsNullOrEmpty(eng))
                    {
                        enginesToExclude.Add(eng);
                    }
                }

                // Now, add the exclusion conditions to the collection
                if (enginesToExclude.Count > 0)
                {
                    // Link the previous condition (U_SerialType) to this new block of conditions
                    con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    for (int i = 0; i < enginesToExclude.Count; i++)
                    {
                        con = oCons.Add();
                        con.Alias = "DistNumber";
                        con.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        con.CondVal = enginesToExclude[i];

                        // Set 'AND' relationship only if it is NOT the last item in the list
                        if (i < enginesToExclude.Count - 1)
                        {
                            con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        }
                    }
                }

                // Apply the conditions to CFL
                oCFL.SetConditions(oCons);
            }


            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region FILTER Key Type
        internal void FilterKeyType(int currentRow)
        {

            _Form.Freeze(true);
            try
            {
                string currentLotNo = _MasterDBDataSource.GetValue("U_LotNo", 0).Trim();
                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_Key");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con;
                //_Cons = new Conditions();

                con = oCons.Add();
                //con.BracketOpenNum = 1;
                con.Alias = "LotNumber";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = currentLotNo;
                //con.BracketCloseNum = 1;
                con.Relationship = BoConditionRelationship.cr_AND;

                con = oCons.Add();
                //con.BracketOpenNum = 1;
                con.Alias = "U_SerialType";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "KEY";
                con.Relationship = BoConditionRelationship.cr_AND;
                //con.BracketCloseNum = 1;

                con = oCons.Add();
                con.Alias = "Quantity";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "1";
                con.Relationship = BoConditionRelationship.cr_AND;

                con = oCons.Add();
                con.Alias = "QuantOut";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "0";

                var enginesToExclude = new System.Collections.Generic.List<string>();
                for (int i = 1; i <= _Matrix.RowCount; i++)
                {
                    if (i == currentRow) continue; // skip current row

                    string eng = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colSetKey").Cells.Item(i).Specific).Value.Trim();
                    if (!string.IsNullOrEmpty(eng))
                    {
                        enginesToExclude.Add(eng);
                    }
                }

                // Now, add the exclusion conditions to the collection
                if (enginesToExclude.Count > 0)
                {
                    // IMPORTANT: Link the previous condition (U_SerialType) to this new block of conditions
                    con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    for (int i = 0; i < enginesToExclude.Count; i++)
                    {
                        con = oCons.Add();
                        con.Alias = "DistNumber";
                        con.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        con.CondVal = enginesToExclude[i];

                        // Set 'AND' relationship only if it is NOT the last item in the list
                        if (i < enginesToExclude.Count - 1)
                        {
                            con.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        }
                    }
                }
                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Filter Model
        internal void FilterModel(int currentRow)
        {

            _Form.Freeze(true);
            try
            {
                string lotNumber = ((SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific).Value.Trim();
                string query = $@"SELECT ""DocEntry"", ""U_Model"" FROM OPDN WHERE ""U_LotNo"" = '{lotNumber}'";
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(query);

                if (!rs.EoF)
                {
                    string docEntry = rs.Fields.Item("DocEntry").Value.ToString();
                    string model1 = rs.Fields.Item("U_Model").Value.ToString();

                    string query1 = $@"SELECT ""ItmsGrpCod"" FROM OITB WHERE ""ItmsGrpNam"" = '{model1}'";
                    SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rs1.DoQuery(query1);
                    if (!rs1.EoF)
                    {
                        itmGrpCode1 = rs1.Fields.Item("ItmsGrpCod").Value.ToString();
                    }
                    else
                    {
                        Utilities.ShowErrorMessage("No Item Group Code found for Model: " + model);
                    }
                }

                SAPbouiCOM.ChooseFromList oCFL = _Form.ChooseFromLists.Item("CFL_Model");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition con;

                //_Cons = new Conditions();
                con = oCons.Add();
                //_Con.BracketOpenNum = 1;
                con.Alias = "ItmsGrpCod";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = itmGrpCode1;
                con.Relationship = BoConditionRelationship.cr_AND;
                //_Con.BracketCloseNum = 1;
                con = oCons.Add();
                con.Alias = "validFor";
                con.Operation = BoConditionOperation.co_EQUAL;
                con.CondVal = "Y";

                _Form.ChooseFromLists.Item("CFL_Model").SetConditions(oCons);

                oCFL.SetConditions(oCons);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            _Form.Freeze(true);
            _Form.DataBrowser.BrowseBy = "txtLotNo";
            _Form.Settings.Enabled = true;

            _Form.Items.Item("txtLotNo").AffectsFormMode = true;
            UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("prodMtrx").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@ENGCHASISMMH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ENGCHASISMMC");


            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific).Value = (Utilities.getMaxColumnValueNum("@ENGCHASISMMH", "U_DocNo"));
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDate").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Open";

            _Form.EnableMenu("1293", true);
            _Form.Freeze(false);



        }
        #endregion

        #region Add New Row
        private void AddNewRow()
        {
            try
            {
                _Form.Freeze(true);

                _Matrix.FlushToDataSource();

                int lastIndex = _ChildDBDataSource.Size - 1;

                // Check if the last row already has data before adding new one
                string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", lastIndex).Trim();
                string chassisNo = _ChildDBDataSource.GetValue("U_ChasisNo", lastIndex).Trim();
                string transNo = _ChildDBDataSource.GetValue("U_TransNo", lastIndex).Trim();
                string setKey = _ChildDBDataSource.GetValue("U_SetKey", lastIndex).Trim();

                if (!string.IsNullOrEmpty(engineNo) ||
                    !string.IsNullOrEmpty(chassisNo) ||
                    !string.IsNullOrEmpty(transNo) ||
                    !string.IsNullOrEmpty(setKey))
                {
                    // Add a new record to child data source
                    _ChildDBDataSource.InsertRecord(_ChildDBDataSource.Size);
                }

                _Matrix.LoadFromDataSource();
                AddSequenceNumbersToMatrix(_Matrix);
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error while adding new row: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Add Seq No
        private void AddSequenceNumbersToMatrix(SAPbouiCOM.Matrix oMatrix)
        {
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@ENGCHASISMMC");

            for (int i = 0; i < _ChildDBDataSource.Size; i++)
            {
                _ChildDBDataSource.SetValue("U_SrNo", i, (i + 1).ToString());
            }

            oMatrix.LoadFromDataSource();
        }
        #endregion

        #region CFL Type
        private string SafeGetValue(SAPbouiCOM.DataTable dt, string col, int row = 0)
        {
            try
            {
                return dt.Columns.Item(col) != null ? dt.GetValue(col, row).ToString().Trim() : "";
            }
            catch { return ""; }
        }
        #endregion

        #region Validation for Matrix
        private bool Validation()
        {
            try
            {
                // Ensure any last-minute user input is pushed from the UI to the datasource.
                _Matrix.FlushToDataSource();

                // --- The Fix: Loop through the DATASOURCE, not the matrix UI ---
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    // Get all values directly from the datasource for the current row 'i'.
                    string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", i).Trim();
                    string chasisNo = _ChildDBDataSource.GetValue("U_ChasisNo", i).Trim();
                    string transNo = _ChildDBDataSource.GetValue("U_TransNo", i).Trim();
                    string setKey = _ChildDBDataSource.GetValue("U_SetKey", i).Trim();
                    string model = _ChildDBDataSource.GetValue("U_ModelCode", i).Trim();

                    // Check if this is the last row in the datasource AND it's completely blank.
                    // If so, it's the "new entry" row and should be ignored.
                    if (i == _ChildDBDataSource.Size - 1 &&
                        string.IsNullOrEmpty(engineNo) && string.IsNullOrEmpty(chasisNo) &&
                        string.IsNullOrEmpty(transNo) && string.IsNullOrEmpty(setKey))
                    {
                        continue; // Skip the blank final row.
                    }

                    // Special check for the very first row being completely blank.
                    if (_ChildDBDataSource.Size <= 1 && string.IsNullOrEmpty(engineNo) && string.IsNullOrEmpty(chasisNo))
                    {
                        Utilities.Message("Matrix cannot be empty. Please enter at least one complete row.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }

                    // If Engine No. has a value, then validate the rest of the row.
                    if (!string.IsNullOrEmpty(engineNo))
                    {
                        // Use "i + 1" in error messages for user-friendly 1-based row numbers.
                        if (string.IsNullOrEmpty(chasisNo))
                        {
                            Utilities.Message($"Chasis No. is required on row {i + 1}.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        if (string.IsNullOrEmpty(transNo))
                        {
                            Utilities.Message($"Transmission No. is required on row {i + 1}.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        if (string.IsNullOrEmpty(setKey))
                        {
                            Utilities.Message($"Key Set is required on row {i + 1}.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        if (string.IsNullOrEmpty(model))
                        {
                            Utilities.Message($"Model Code is required on row {i + 1}.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                    }
                    // Check for partially filled rows (e.g., Chasis No. exists but Engine No. doesn't).
                    else if (!string.IsNullOrEmpty(chasisNo) || !string.IsNullOrEmpty(transNo) || !string.IsNullOrEmpty(setKey))
                    {
                        Utilities.Message($"Engine No. must be filled on row {i + 1} because other fields have data.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }

                return true; // All rows are valid.
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("An error occurred during matrix validation: " + ex.Message);
                return false;
            }
        }

        #endregion

        #region Disable Form Item
        public void Disable()
        {
            _Form.Freeze(true);
            try
            {
                _Form.Items.Item("txtLotNo").Enabled = false;
                _Form.Items.Item("btnProdO").Enabled = true;
                SAPbouiCOM.Matrix mtx = (SAPbouiCOM.Matrix)_Form.Items.Item("prodMtrx").Specific;

                // Get column indexes dynamically
                int colEngine = GetColumnIndex(mtx, "colEngine");
                int colChasis = GetColumnIndex(mtx, "colChasis");
                int colTrans = GetColumnIndex(mtx, "colTrans");
                int colSetKey = GetColumnIndex(mtx, "colSetKey");
                int colModel = GetColumnIndex(mtx, "colModel");

                int rowCount = _ChildDBDataSource.Size;

                for (int i = 0; i < rowCount; i++)
                {
                    int row = i + 1;

                    string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", i).Trim();
                    string chasisNo = _ChildDBDataSource.GetValue("U_ChasisNo", i).Trim();
                    string transNo = _ChildDBDataSource.GetValue("U_TransNo", i).Trim();
                    string setKey = _ChildDBDataSource.GetValue("U_setKey", i).Trim();
                    string prodEnt = _ChildDBDataSource.GetValue("U_ProdOrdNo", i).Trim();
                    string modelCode = _ChildDBDataSource.GetValue("U_ModelCode", i).Trim();

                    bool isLastRow = string.IsNullOrEmpty(engineNo) && string.IsNullOrEmpty(chasisNo) && string.IsNullOrEmpty(transNo) && string.IsNullOrEmpty(setKey) && string.IsNullOrEmpty(prodEnt) && string.IsNullOrEmpty(modelCode);
                    if (isLastRow)
                        continue;

                    if (!string.IsNullOrEmpty(prodEnt))
                    {
                        mtx.CommonSetting.SetCellEditable(row, colEngine, false);
                        mtx.CommonSetting.SetCellEditable(row, colChasis, false);
                        mtx.CommonSetting.SetCellEditable(row, colTrans, false);
                        mtx.CommonSetting.SetCellEditable(row, colSetKey, false);
                        mtx.CommonSetting.SetCellEditable(row, colModel, false);
                        continue;
                    }

                    mtx.CommonSetting.SetCellEditable(row, colEngine, string.IsNullOrEmpty(engineNo));
                    mtx.CommonSetting.SetCellEditable(row, colChasis, string.IsNullOrEmpty(chasisNo));
                    mtx.CommonSetting.SetCellEditable(row, colTrans, string.IsNullOrEmpty(transNo));
                    mtx.CommonSetting.SetCellEditable(row, colSetKey, string.IsNullOrEmpty(setKey));
                    mtx.CommonSetting.SetCellEditable(row, colModel, string.IsNullOrEmpty(modelCode));
                }


            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText(ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                _Form.Freeze(false);
            }

        }
        #endregion

        #region Get Index of a column
        private int GetColumnIndex(SAPbouiCOM.Matrix mtx, string columnUid)
        {
            for (int i = 1; i <= mtx.Columns.Count; i++)
            {
                if (mtx.Columns.Item(i).UniqueID == columnUid)
                    return i;
            }
            return -1;
        }
        #endregion

        //#region PIN Code Generation
        //private void GenerateVINPinFromMatrix()
        //{
        //    SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;

        //    SAPbouiCOM.Matrix mtxSrc = (SAPbouiCOM.Matrix)oForm.Items.Item("prodMtrx").Specific;
        //    SAPbouiCOM.Matrix mtxDes = (SAPbouiCOM.Matrix)oForm.Items.Item("mtxGenPin").Specific;

        //    SAPbouiCOM.DBDataSource dsDes = oForm.DataSources.DBDataSources.Item("@ENGCHASISMMC1");
        //    SAPbouiCOM.DBDataSource dsSrc = oForm.DataSources.DBDataSources.Item("@ENGCHASISMMC");

        //    int[] prime = { 3, 5, 7, 11, 13, 17 };

        //    oForm.Freeze(true);

        //    mtxSrc.FlushToDataSource();
        //    mtxDes.FlushToDataSource();
        //    if (dsDes.Size > 0)
        //    {
        //        string chk = dsDes.GetValue("U_VIN", 0).Trim();
        //        if (string.IsNullOrEmpty(chk))
        //            dsDes.RemoveRecord(0);
        //    }

        //    for (int i = 1; i <= mtxSrc.VisualRowCount; i++)
        //    {
        //        string chassis = ((SAPbouiCOM.EditText)mtxSrc.Columns.Item("colChasis").Cells.Item(i).Specific).Value.Trim();
        //        if (string.IsNullOrEmpty(chassis)) continue;

        //        // Extract last 6 digits from full chassis
        //        string last6 = new string(chassis.Where(char.IsDigit).ToArray()).Substring(chassis.Where(char.IsDigit).Count() - 6, 6);

        //        int idx = dsDes.Size;

        //        if (idx == 0)
        //            dsDes.InsertRecord(0);
        //        else
        //            dsDes.InsertRecord(idx - 1);

        //        idx = dsDes.Size - 1;
        //        //-------------------------------------------------------------

        //        dsDes.SetValue("U_VIN", idx, chassis);
        //        dsDes.SetValue("U_Last6", idx, last6);

        //        int totalMul = 0;

        //        // Digit split + prime multiplication 
        //        for (int j = 0; j < last6.Length; j++)
        //        {
        //            int d = Convert.ToInt32(last6[j].ToString());

        //            // Store digits
        //            dsDes.SetValue($"U_D{j + 1}", idx, d.ToString());

        //            // Prime multiplication
        //            int mul = d * prime[j];
        //            totalMul += mul;
        //        }

        //        //int primeNo = 89;
        //        //// Generate PIN - Try 89 first
        //        //int multiple = Convert.ToInt32(last6) * primeNo;
        //        //int addition = multiple + totalMul;
        //        //int Modulous = addition % 1000000;

        //        //// If leading zero occurs -> recalculate using 79
        //        //if (Modulous.ToString().Length < 6)
        //        //{
        //        //    //set dynamic value
        //        //    primeNo = 83;
        //        //    multiple = Convert.ToInt32(last6) * primeNo;
        //        //    addition = multiple + totalMul;
        //        //    Modulous = addition % 1000000;
        //        //}
        //        int[] primeList = { 89, 83, 79, 73, 71 };

        //        int primeNo = 89;
        //        int multiple = 0;
        //        int addition = 0;
        //        int Modulous = 0;

        //        foreach (int p in primeList)
        //        {
        //            primeNo = p;

        //            multiple = Convert.ToInt32(last6) * primeNo;
        //            addition = multiple + totalMul;
        //            Modulous = addition % 1000000;

        //            if (Modulous.ToString("D6")[0] != '0')
        //                break;
        //        }

        //        // Save values
        //        dsDes.SetValue("U_WSum", idx, totalMul.ToString());
        //        dsDes.SetValue("U_Mult", idx, multiple.ToString());
        //        dsDes.SetValue("U_AddVal", idx, addition.ToString());
        //        dsDes.SetValue("U_ModVal", idx, Modulous.ToString());
        //        dsDes.SetValue("U_PinCode", idx, Modulous.ToString("D6"));  // Always keep 6 digits
        //        dsDes.SetValue("U_CalcPrime", idx, primeNo.ToString());
        //        dsSrc.SetValue("U_PinCode", idx, Modulous.ToString("D6"));
        //    }

        //    mtxDes.LoadFromDataSource();
        //    mtxSrc.LoadFromDataSource();

        //    this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
        //    oForm.Freeze(false);
        //}
        //#endregion

        #region PIN Code Generation
        private void GenerateVINPinFromMatrix()
        {
            SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;

            SAPbouiCOM.Matrix mtxSrc = (SAPbouiCOM.Matrix)oForm.Items.Item("prodMtrx").Specific;
            SAPbouiCOM.Matrix mtxDes = (SAPbouiCOM.Matrix)oForm.Items.Item("mtxGenPin").Specific;

            SAPbouiCOM.DBDataSource dsSrc = oForm.DataSources.DBDataSources.Item("@ENGCHASISMMC");
            SAPbouiCOM.DBDataSource dsDes = oForm.DataSources.DBDataSources.Item("@ENGCHASISMMC1");

            int[] prime = { 3, 5, 7, 11, 13, 17 };
            int[] primeList = { 89, 83, 79, 73, 71, 67, 61, 59, 53 };

            oForm.Freeze(true);
            try
            {
                mtxSrc.FlushToDataSource();
                mtxDes.FlushToDataSource();

                // Remove empty first row in destination if it exists and has no VIN
                if (dsDes.Size > 0)
                {
                    string chk = dsDes.GetValue("U_VIN", 0).Trim();
                    if (string.IsNullOrEmpty(chk))
                        dsDes.RemoveRecord(0);
                }

                // Iterate through the Source Matrix Rows
                for (int i = 0; i < dsSrc.Size; i++)
                {
                    string chassis = dsSrc.GetValue("U_ChasisNo", i).Trim();
                    string existingPin = dsSrc.GetValue("U_PinCode", i).Trim();

                    // SKIP if Chassis is empty (usually the last row)
                    if (string.IsNullOrEmpty(chassis)) continue;

                    // SKIP if PIN is ALREADY generated in the source
                    if (!string.IsNullOrEmpty(existingPin)) continue;

                    bool foundInDestination = false;
                    string existingCalculatedPin = "";

                    //if chasis already exists in Destination DB, prevents adding a duplicate row to the PIN table
                    for (int k = 0; k < dsDes.Size; k++)
                    {
                        string desVin = dsDes.GetValue("U_VIN", k).Trim();
                        if (desVin == chassis)
                        {
                            foundInDestination = true;
                            existingCalculatedPin = dsDes.GetValue("U_PinCode", k).Trim();
                            break;
                        }
                    }

                    if (foundInDestination)
                    {
                        // If it exists in the PIN table but was missing in the Source,
                        // just copy the PIN back to the Source and DO NOT add a new row to dsDes.
                        if (!string.IsNullOrEmpty(existingCalculatedPin))
                        {
                            dsSrc.SetValue("U_PinCode", i, existingCalculatedPin);
                        }
                        continue; // Move to next row, do not run calculation logic
                    }

                    // Extract last 6 digits
                    string last6 = new string(chassis.Where(char.IsDigit).ToArray());
                    if (last6.Length >= 6)
                        last6 = last6.Substring(last6.Length - 6, 6);
                    else
                        last6 = last6.PadLeft(6, '0');

                    int idx = dsDes.Size;
                    dsDes.InsertRecord(idx);


                    idx = dsDes.Size - 1;

                    dsDes.SetValue("U_SerialNo", idx, (idx + 1).ToString());

                    dsDes.SetValue("U_VIN", idx, chassis);
                    dsDes.SetValue("U_Last6", idx, last6);

                    int totalMul = 0;

                    // Digit split + prime multiplication 
                    for (int j = 0; j < last6.Length; j++)
                    {
                        if (j < prime.Length)
                        {
                            //store digits
                            int d = Convert.ToInt32(last6[j].ToString());
                            dsDes.SetValue($"U_D{j + 1}", idx, d.ToString());

                            // Prime multiplication
                            int mul = d * prime[j];
                            totalMul += mul;
                        }
                    }

                    // Determine final PIN using the Prime List strategy
                    int primeNo = 89;
                    int multiple = 0;
                    int addition = 0;
                    int Modulous = 0;

                    foreach (int p in primeList)
                    {
                        primeNo = p;
                        multiple = Convert.ToInt32(last6) * primeNo;
                        addition = multiple + totalMul;
                        Modulous = addition % 1000000;

                        // Check if result has leading zero (less than 6 digits)
                        if (Modulous.ToString("D6")[0] != '0')
                            break;
                    }

                    // Save Calculated Values to Destination
                    dsDes.SetValue("U_WSum", idx, totalMul.ToString());
                    dsDes.SetValue("U_Mult", idx, multiple.ToString());
                    dsDes.SetValue("U_AddVal", idx, addition.ToString());
                    dsDes.SetValue("U_ModVal", idx, Modulous.ToString());
                    dsDes.SetValue("U_PinCode", idx, Modulous.ToString("D6"));
                    dsDes.SetValue("U_CalcPrime", idx, primeNo.ToString());

                    // Save Calculated PIN back to Source
                    dsSrc.SetValue("U_PinCode", i, Modulous.ToString("D6"));
                }

                // Reload matrices to show new data
                mtxDes.LoadFromDataSource();
                mtxSrc.LoadFromDataSource();

                //if (oForm.Mode == BoFormMode.fm_OK_MODE)
                //{
                //    oForm.Mode = BoFormMode.fm_UPDATE_MODE;
                //}
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error generating PIN: " + ex.Message);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        #endregion

        #region delete rows and resequence
        //private void DeleteRowAndResequence()
        //{
        //    try
        //    {
        //        _Form.Freeze(true);

        //        _Matrix.FlushToDataSource();
        //        int visualRowIndex = -1;

        //        int selRow = _Matrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

        //        if (selRow > 0)
        //        {
        //            visualRowIndex = selRow;
        //        }
        //        else if (_rowToDelete > 0)
        //        {
        //            visualRowIndex = _rowToDelete;
        //        }

        //        if (visualRowIndex <= 0)
        //        {
        //            Utilities.Message("No row selected for deletion.", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        //            return;
        //        }

        //        int dsIndex = visualRowIndex - 1; 

        //        if (dsIndex >= _ChildDBDataSource.Size)
        //        {
        //            return;
        //        }

        //        string prodOrdNo = _ChildDBDataSource.GetValue("U_ProdOrdNo", dsIndex).Trim();
        //        if (!string.IsNullOrEmpty(prodOrdNo))
        //        {
        //            Utilities.Application.SBO_Application.MessageBox("Cannot delete this row because a Production Order is linked.");
        //            return;
        //        }

        //        // Get the chassis number before deleting the row
        //        string chasisNo = _ChildDBDataSource.GetValue("U_ChasisNo", dsIndex).Trim();

        //        //delete from pincode tab
        //        if (!string.IsNullOrEmpty(chasisNo))
        //        {
        //            SAPbouiCOM.DBDataSource dsPin = _Form.DataSources.DBDataSources.Item("@ENGCHASISMMC1");
        //            SAPbouiCOM.Matrix mtxPin = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxGenPin").Specific;

        //            for (int p = 0; p < dsPin.Size; p++)
        //            {
        //                // matching Chassis in the Pin Code 
        //                if (dsPin.GetValue("U_VIN", p).Trim() == chasisNo)
        //                {
        //                    dsPin.RemoveRecord(p);
        //                    mtxPin.LoadFromDataSource(); 
        //                    break;
        //                }
        //            }
        //        }

        //        // Remove the record from the main Engine/Chassis datasource
        //        _ChildDBDataSource.RemoveRecord(dsIndex);

        //        // Update OSRN table
        //        if (!string.IsNullOrEmpty(chasisNo))
        //        {
        //            try
        //            {
        //                rsUpdate = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                string updateQuery = $@"UPDATE ""OSRN"" 
        //                                SET ""U_PinCode"" = NULL, ""U_EngChasisMMNo"" = NULL 
        //                                WHERE ""DistNumber"" = '{chasisNo}'";
        //                rsUpdate.DoQuery(updateQuery);
        //            }
        //            catch (Exception ex)
        //            {
        //                Utilities.Application.SBO_Application.StatusBar.SetText("Row deleted, but failed to update OSRN: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        //            }
        //        }

        //        //Resequence the SR numbers and update quantities
        //        int validRowsCount = 0;
        //        for (int i = 0; i < _ChildDBDataSource.Size; i++)
        //        {
        //            _ChildDBDataSource.SetValue("U_SrNo", i, (i + 1).ToString());

        //            string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", i).Trim();
        //            if (!string.IsNullOrEmpty(engineNo))
        //            {
        //                validRowsCount++;
        //            }
        //        }

        //        _MasterDBDataSource.SetValue("U_CreatedQty", 0, validRowsCount.ToString());

        //        double toCreateQty = 0;
        //        string toCreateStr = _MasterDBDataSource.GetValue("U_ToCreate", 0).Trim();
        //        if (double.TryParse(toCreateStr, out double parsedQty))
        //        {
        //            toCreateQty = parsedQty;
        //        }

        //        double newRemainingQty = toCreateQty - validRowsCount;
        //        _MasterDBDataSource.SetValue("U_RemQty", 0, newRemainingQty.ToString());

        //        _Matrix.LoadFromDataSource();
        //        _rowToDelete = -1;

        //        if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
        //        {
        //            _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        //        }

        //        _Form.Items.Item("1").Click();
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.ShowErrorMessage("Error deleting row: " + ex.Message);
        //    }
        //    finally
        //    {
        //        _Form.Freeze(false);
        //    }
        //}

        private void DeleteRowAndResequence()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                int selRow = _Matrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);
                int visualRowIndex = selRow > 0 ? selRow : _rowToDelete;

                if (visualRowIndex <= 0) return;

                int dsIndex = visualRowIndex - 1;
                if (dsIndex >= _ChildDBDataSource.Size) return;

                string prodOrdNo = _ChildDBDataSource.GetValue("U_ProdOrdNo", dsIndex).Trim();
                if (!string.IsNullOrEmpty(prodOrdNo))
                {
                    Utilities.Application.SBO_Application.MessageBox("Cannot delete this row because a Production Order is linked.");
                    return;
                }

                string chasisNo = _ChildDBDataSource.GetValue("U_ChasisNo", dsIndex).Trim();

                // 1. Handle PIN Table Deletion and Resequencing
                if (!string.IsNullOrEmpty(chasisNo))
                {
                    SAPbouiCOM.DBDataSource dsPin = _Form.DataSources.DBDataSources.Item("@ENGCHASISMMC1");
                    SAPbouiCOM.Matrix mtxPin = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxGenPin").Specific;

                    bool pinDeleted = false;
                    for (int p = 0; p < dsPin.Size; p++)
                    {
                        if (dsPin.GetValue("U_VIN", p).Trim() == chasisNo)
                        {
                            dsPin.RemoveRecord(p);
                            pinDeleted = true;
                            break;
                        }
                    }

                    // --- NEW: RESEQUENCE PIN TABLE SERIAL NUMBERS ---
                    if (pinDeleted)
                    {
                        for (int k = 0; k < dsPin.Size; k++)
                        {
                            dsPin.SetValue("U_SerialNo", k, (k + 1).ToString());
                        }
                        mtxPin.LoadFromDataSource();
                    }
                }

                // 2. Main Row Deletion
                _ChildDBDataSource.RemoveRecord(dsIndex);

                // 3. Update System Serial Table (OSRN)
                if (!string.IsNullOrEmpty(chasisNo))
                {
                    try
                    {
                        SAPbobsCOM.Recordset rsUpdateLocal = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string updateQuery = $@"UPDATE ""OSRN"" SET ""U_PinCode"" = NULL, ""U_EngChasisMMNo"" = NULL WHERE ""DistNumber"" = '{chasisNo}'";
                        rsUpdateLocal.DoQuery(updateQuery);
                    }
                    catch { /* Log if necessary */ }
                }

                // 4. Resequence Main Matrix SR Numbers & Update Quantities
                int validRowsCount = 0;
                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    _ChildDBDataSource.SetValue("U_SrNo", i, (i + 1).ToString());
                    if (!string.IsNullOrEmpty(_ChildDBDataSource.GetValue("U_EngineNo", i).Trim()))
                        validRowsCount++;
                }

                _MasterDBDataSource.SetValue("U_CreatedQty", 0, validRowsCount.ToString());
                double.TryParse(_MasterDBDataSource.GetValue("U_ToCreate", 0).Trim(), out double toCreateQty);
                _MasterDBDataSource.SetValue("U_RemQty", 0, (toCreateQty - validRowsCount).ToString());

                _Matrix.LoadFromDataSource();
                _rowToDelete = -1;

                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                _Form.Items.Item("1").Click();
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error during deletion: " + ex.Message);
            }
            finally { _Form.Freeze(false); }
        }

        #endregion

        #region Right Click Event
        public override void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo eventInfo, ref bool BubbleEvent)
        {
            base.RightClick_Event(ref eventInfo, ref BubbleEvent);

            if (eventInfo.BeforeAction)
            {
                if (eventInfo.ItemUID == "prodMtrx")
                {
                    _rowToDelete = eventInfo.Row;
                    try
                    {
                        if (_rowToDelete > 0)
                        {
                            _Matrix.SelectRow(_rowToDelete, true, false);
                        }
                    }
                    catch { /* Ignore error if selection fails */ }
                }
                else
                {
                    _rowToDelete = -1; // Reset if clicked elsewhere
                }
            }
        }
        #endregion

        #region Update Push Inventory Text
        private void UpdatePushToInventoryDisplay()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                for (int i = 0; i < _ChildDBDataSource.Size; i++)
                {
                    string status = _ChildDBDataSource.GetValue("U_Status", i).Trim();

                    if (status == "Completed")
                    {
                        _ChildDBDataSource.SetValue("U_PushInvt", i, "Push to Inventory");
                    }
                    else
                    {
                        _ChildDBDataSource.SetValue("U_PushInvt", i, "");
                    }
                }

                _Matrix.LoadFromDataSource();

                string colUID = "colInvt";
                int linkBlueColor = 16711680;// colour blue

                for (int i = 1; i <= _Matrix.VisualRowCount; i++)
                {
                    SAPbouiCOM.Cell oCell = _Matrix.Columns.Item(colUID).Cells.Item(i);
                    SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oCell.Specific;

                    if (oEdit.Value.Trim() == "Push to Inventory")
                    {
                        oEdit.ForeColor = linkBlueColor;
                        oEdit.TextStyle = (int)BoTextStyle.ts_UNDERLINE;
                        oEdit.FontSize = 11;
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error updating Push Inventory text: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Ensure Rows Match To Create Quantity
        private void EnsureRowsForRequiredQty()
        {
            try
            {
                _Form.Freeze(true);
                _Matrix.FlushToDataSource();

                //Get Required Quantity
                string toCreateStr = _MasterDBDataSource.GetValue("U_ToCreate", 0).Trim();
                double.TryParse(toCreateStr, out double requiredQty);

                // Get Current Row Count
                int currentRows = _ChildDBDataSource.Size;

                //Only add a row if it is below the Required Quantity
                if (currentRows < requiredQty)
                {
                    // Check if the matrix is completely empty
                    if (currentRows == 0)
                    {
                        _ChildDBDataSource.InsertRecord(0);
                        _ChildDBDataSource.SetValue("U_SrNo", 0, "1");
                    }
                    else
                    {
                        // Check if the LAST row is filled
                        int lastIndex = currentRows - 1;
                        string engineNo = _ChildDBDataSource.GetValue("U_EngineNo", lastIndex).Trim();
                        string chasisNo = _ChildDBDataSource.GetValue("U_ChasisNo", lastIndex).Trim();
                        string model = _ChildDBDataSource.GetValue("U_ModelCode", lastIndex).Trim();

                        // If the last row has data, add ONE new empty row
                        if (!string.IsNullOrEmpty(engineNo) || !string.IsNullOrEmpty(chasisNo) || !string.IsNullOrEmpty(model))
                        {
                            _ChildDBDataSource.InsertRecord(currentRows);
                            _ChildDBDataSource.SetValue("U_SrNo", currentRows, (currentRows + 1).ToString());

                            // Clear fields to ensure it is clean
                            _ChildDBDataSource.SetValue("U_EngineNo", currentRows, "");
                            _ChildDBDataSource.SetValue("U_ChasisNo", currentRows, "");
                        }
                    }

                    _Matrix.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error ensuring row: " + ex.Message);
            }
            finally
            {
                _Form.Freeze(false);
            }
        }
        #endregion

        #region Move to next 
        private void MoveFocusToNext(string nextColUID, int row)
        {
            System.Threading.Tasks.Task.Delay(150).ContinueWith(t =>
            {
                try
                {
                    SAPbouiCOM.Cell nextCell = _Matrix.Columns.Item(nextColUID).Cells.Item(row);
                    nextCell.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    ((SAPbouiCOM.EditText)nextCell.Specific).Active = true;
                }
                catch { }
            });
        }
        //MoveFocusToNext("colEngine", pVal.Row + 1);
        #endregion

        private int Create_ReceiptFrom_Production(int prodOrderDocEntry, string chasis, string engine, string key, string transm)
        {
            int newDocEntry = 0;
            try
            {
                // 1. Get Parent Item details from Production Order Header (OWOR)
                SAPbobsCOM.Recordset rsOrder = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // Fetch ItemCode and Warehouse from the specific Prod Order
                rsOrder.DoQuery($@"SELECT T0.""ItemCode"", T0.""Warehouse"" FROM OWOR T0 WHERE T0.""DocEntry"" = {prodOrderDocEntry}");

                if (rsOrder.EoF) throw new Exception("Production Order details not found.");

                string parentItem = rsOrder.Fields.Item("ItemCode").Value.ToString();
                string whs = rsOrder.Fields.Item("Warehouse").Value.ToString();
                double qtyToReceipt = 1; // We are pushing 1 unit per row

                // 2. Initialize Inventory Gen Entry (Receipt from Production)
                SAPbobsCOM.Documents oReceipt = (SAPbobsCOM.Documents)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

                oReceipt.DocDate = DateTime.Now;
                oReceipt.Comments = $"Auto Receipt for Chassis: {chasis} | Engine: {engine}";
                oReceipt.UserFields.Fields.Item("U_ProdEntry").Value = prodOrderDocEntry;

                // 3. Add Line Item
                oReceipt.Lines.BaseType = 202;            // 202 = Production Order
                oReceipt.Lines.BaseEntry = prodOrderDocEntry;
                oReceipt.Lines.WarehouseCode = whs;
                oReceipt.Lines.Quantity = qtyToReceipt;
                // Note: For Receipts from Production, BaseLine is usually not required if there is only one parent item, 
                // but explicit TransactionType can be safer.

                // 4. Handle Serial Numbers
                // Check if Item is Serial Managed
                SAPbobsCOM.Recordset rsItem = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsItem.DoQuery($"SELECT \"ManSerNum\" FROM OITM WHERE \"ItemCode\" = '{parentItem}'");

                SAPbobsCOM.Recordset rsEngPO = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsEngPO.DoQuery($@"SELECT T0.""U_LotNo"", T0.""U_EnChNo"", T0.""U_PinCode"" FROM ""@ENGCHASPO"" T0 WHERE T0.""U_ProdDocEntry"" = {prodOrderDocEntry}");

                Recordset rsOSRN = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                rsOSRN.DoQuery($@"SELECT ""U_ModelYear"" FROM ""@ENGCHASISMMC"" WHERE ""U_ChasisNo"" = '{chasis}'");

                if (rsItem.RecordCount > 0 && rsItem.Fields.Item("ManSerNum").Value.ToString() == "Y")
                {
                    // Logic from ProductionPlanningLMC:
                    // InternalSerialNumber = Engine
                    // ManufacturerSerialNumber = Chassis
                    string LotNo = rsEngPO.Fields.Item("U_LotNo").Value.ToString();
                    string EngineChasisNo = rsEngPO.Fields.Item("U_EnChNo").Value.ToString();
                    string PinCode = rsEngPO.Fields.Item("U_PinCode").Value.ToString();
                    string modelYear = rsOSRN.Fields.Item("U_ModelYear").Value.ToString();


                    oReceipt.Lines.SerialNumbers.InternalSerialNumber = engine ?? "";
                    oReceipt.Lines.SerialNumbers.ManufacturerSerialNumber = chasis ?? "";
                    oReceipt.Lines.SerialNumbers.BatchID = LotNo ?? "";
                    oReceipt.Lines.SerialNumbers.ManufactureDate = DateTime.Now; 

                    // Set User Fields on Serial Number (Ensure these UDFs exist on OSRI/SRI1 in SAP)
                    try { oReceipt.Lines.SerialNumbers.UserFields.Fields.Item("U_TransmissionNo").Value = transm ?? ""; } catch { }
                    try { oReceipt.Lines.SerialNumbers.UserFields.Fields.Item("U_KeyNo").Value = key ?? ""; } catch { }
                    try { oReceipt.Lines.SerialNumbers.UserFields.Fields.Item("U_PinCode").Value = PinCode ?? ""; } catch { }
                    try { oReceipt.Lines.SerialNumbers.UserFields.Fields.Item("U_EngChasisMMNo").Value = EngineChasisNo ?? ""; } catch { }
                    try { oReceipt.Lines.SerialNumbers.UserFields.Fields.Item("U_ModelYear").Value = modelYear ?? ""; } catch { }

                    oReceipt.Lines.SerialNumbers.Quantity = 1;
                    oReceipt.Lines.SerialNumbers.Add();
                }

                oReceipt.Lines.Add();

                // 5. Add Document
                int ret = oReceipt.Add();
                if (ret != 0)
                {
                    //Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                    //throw new Exception($"Receipt creation failed: {errMsg}");
                    string diErr = Utilities.Application.Company.GetLastErrorDescription();
                    int diCode = Utilities.Application.Company.GetLastErrorCode();
                    throw new Exception($"Receipt from Production DI Error {diCode}: {diErr}");
                }
                else
                {
                    string tempKey = Utilities.Application.Company.GetNewObjectKey();
                    if (int.TryParse(tempKey, out newDocEntry))
                    {
                        return newDocEntry;
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText(
                    ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return newDocEntry;
            }
            finally
            {
                // Cleanup logic if needed
            }
            return 0;
        }

        private void UpdateRowAfterReceipt(string prodDocEntry, string chassisNo, int receiptDocEntry)
        {
            SAPbobsCOM.Recordset rs = null;
            try
            {
                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Update Table (@ENGCHASPO) 
                // Set Status to Manufactured and save the Receipt DocEntry
                string queryChild = $@"UPDATE ""@ENGCHASPO"" 
                               SET ""U_Status"" = 'Manufactured', 
                                   ""U_ReceiptFrmProd"" = '{receiptDocEntry}'
                               WHERE ""U_ProdDocEntry"" = '{prodDocEntry}' AND ""U_ChasisNo"" = '{chassisNo}'";
                rs.DoQuery(queryChild);
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error updating tables: " + ex.Message);
            }
            finally
            {
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
    }
}
