using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Production_Planning_LMC
{
    class ClsJobOrderExecution : Base
    {
        #region VARIABLE DECALARTION
        public SAPbouiCOM.Matrix _Matrix;
        public SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        string selectedLotNo;
        private SAPbouiCOM.Conditions _Cons;
        private SAPbouiCOM.Condition _Con;
        string productionNo;
        private bool _isHandlingCFLEvent = false;
        private SAPbouiCOM.BoFormMode _lastMode;
        #endregion

        #region CONSTRUCTOR and DISTRUCTOR
        public ClsJobOrderExecution() : base()
        {

        }
        ~ClsJobOrderExecution()
        {

        }
        #endregion


        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            // After the data is loaded into the form
            if (!BusinessObjectInfo.BeforeAction && BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD)
            {
                try
                {
                    _Form.Freeze(true);

                    // Get status from the Master DBDataSource
                    // Assuming index 0 as it is a header
                    string currentStatus = _Form.DataSources.DBDataSources.Item("@WJOBEXEH").GetValue("U_Status", 0).Trim();

                    // Call the UI state controller
                    SetRemarkFieldsState(currentStatus);
                    LoadBreakCategoryValues();
                }
                catch (Exception ex)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Error on Data Load: " + ex.Message);
                }
                finally
                {
                    _Form.Freeze(false);
                }
            }
        }

        #region ItemEvent
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            //#region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" &&  _Form.Mode == BoFormMode.fm_UPDATE_MODE)
                        {
                            _lastMode = _Form.Mode;
                        }

                        break;

            //        case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

            //        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
            //            if (pVal.ItemUID == "txtLotNo")
            //            {
            //                this.FilterItemType("");
            //            }
            //            if (pVal.ItemUID == "txtPOEnt")
            //            {
            //                this.FilterProducNo("");
            //            }


            //            break;
                }
            }

            //#endregion

            //#region After Action
            else
            {
               switch (pVal.EventType)
            {
            //        #region Set CFL Values
                   

            //        case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
            //            _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

            //            if (pVal.ItemUID == "txtPOEnt" && _SysCFLEvent.SelectedObjects != null)
            //            {
            //                if (_isHandlingCFLEvent)
            //                {
            //                    return; // Exit if we are already in this handler
            //                }

            //                _isHandlingCFLEvent = true;
            //                try
            //                {
            //                    _Form.Freeze(true);

            //                    productionNo = _SysCFLEvent.SelectedObjects.GetValue("DocEntry", 0).ToString().Trim();
            //                    _MasterDBDataSource.SetValue("U_PrdOrdEnt", 0, productionNo);

            //                    SAPbobsCOM.Recordset rsLot = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                    string queryLot = $@"SELECT COUNT(*) AS ""CNT"", MAX(""DocEntry"") AS ""DocEntry"" FROM ""@WRKORDRDTLSH"" WHERE ""U_PrdOrdEnt"" = '{productionNo.Replace("'", "''")}'";
            //                    rsLot.DoQuery(queryLot);
            //                    int count = Convert.ToInt32(rsLot.Fields.Item("CNT").Value);

            //                    if (count == 0)
            //                    {

            //                        // Fetch Production Order Number and Product Details
            //                        //string query = $@"SELECT ""DocEntry"", ""DocNum"", ""ItemCode"", ""ProdName"" FROM OWOR WHERE ""U_LotNo"" = '{selectedLotNo}'";
            //                        string query = $@"SELECT ""DocNum"", ""ItemCode"", ""ProdName"", ""U_LotNo"", ""PlannedQty"" FROM OWOR WHERE ""DocEntry"" = '{productionNo}'";
            //                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                        rs.DoQuery(query);

            //                        if (!rs.EoF)
            //                        {
            //                            //string docEntry = rs.Fields.Item("DocEntry").Value.ToString();
            //                            string docNum = rs.Fields.Item("DocNum").Value.ToString();
            //                            string itemCode = rs.Fields.Item("ItemCode").Value.ToString();
            //                            string itemName = rs.Fields.Item("ProdName").Value.ToString();
            //                            string lotNo = rs.Fields.Item("U_LotNo").Value.ToString();
            //                            string plannedQty = rs.Fields.Item("PlannedQty").Value.ToString();
            //                            //_MasterDBDataSource.SetValue("U_ProdOrdEntry", 0, docEntry);
            //                            _MasterDBDataSource.SetValue("U_ProdOrdNo", 0, docNum);
            //                            _MasterDBDataSource.SetValue("U_ProductCode", 0, itemCode);
            //                            _MasterDBDataSource.SetValue("U_ProductName", 0, itemName);
            //                            _MasterDBDataSource.SetValue("U_LotNo", 0, lotNo);
            //                            _MasterDBDataSource.SetValue("U_PlannedQty", 0, plannedQty);

            //                            string modelQuery = $@"Select T0.""ItmsGrpNam"" from OITB T0 inner join OITM T1 on T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" where T1.""ItemCode"" = '{itemCode}' ";
            //                            SAPbobsCOM.Recordset rsModel = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                            rsModel.DoQuery(modelQuery);

            //                            string model = rsModel.Fields.Item("ItmsGrpNam").Value.ToString();
            //                            _MasterDBDataSource.SetValue("U_Model", 0, model);

            //                            //fetch Doc No of Engine/Chasis Mapping Master
            //                            string query1 = $@"SELECT ""U_DocNo"" FROM ""@ENGCHASISMMH"" WHERE ""U_LotNo"" = '{lotNo}'";
            //                            SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                            rs1.DoQuery(query1);
            //                            if (!rs1.EoF)
            //                            {
            //                                string docNo = rs1.Fields.Item("U_DocNo").Value.ToString();
            //                                _MasterDBDataSource.SetValue("U_AdvancedSBNo", 0, docNo);
            //                            }
            //                            else
            //                            {
            //                                Utilities.ShowErrorMessage("No Engine/Chasis Mapping Master found for Lot No: " + selectedLotNo);
            //                            }
            //                        }
            //                        else
            //                        {
            //                            Utilities.ShowErrorMessage("No Production Order found for Lot No: " + selectedLotNo);
            //                        }

            //                        //// Prepare Matrix for user input
            //                        //_Matrix.Clear();
            //                        //_ChildDBDataSource.Clear();
            //                        //_ChildDBDataSource.InsertRecord(0);

            //                        //_Matrix.LoadFromDataSource();
            //                        //AddSequenceNumbersToMatrix(_Matrix);
            //                    }
            //                    else if (count == 1)
            //                    {
            //                        SAPbobsCOM.Recordset rsLot1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                        string codeQuery1 = $@"SELECT ""U_PrdOrdEnt""
            //                         FROM ""@WRKORDRDTLSH"" 
            //                         WHERE ""U_PrdOrdEnt"" = '{productionNo.Replace("'", "''")}'";

            //                        rsLot1.DoQuery(codeQuery1);
            //                        _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

            //                        string recordLotNo = rsLot1.Fields.Item("U_PrdOrdEnt").Value.ToString();

            //                        ((SAPbouiCOM.EditText)_Form.Items.Item("txtPOEnt").Specific).Value = recordLotNo;
            //                        _MasterDBDataSource.SetValue("U_PrdOrdEnt", 0, recordLotNo);

            //                        System.Threading.Thread.Sleep(60);

            //                        _Form.Items.Item("1").Click();

            //                        System.Threading.Thread.Sleep(60);
            //                    }
            //                }
            //                catch (Exception ex)
            //                {
            //                    Utilities.ShowErrorMessage("Error while setting LotNo CFL value: " + ex.Message);
            //                }
            //                finally
            //                {
            //                    _Form.Freeze(false);

            //                    _isHandlingCFLEvent = false;
            //                }
                            
            //            }

            //            break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if(pVal.ItemUID == "1" && _lastMode == BoFormMode.fm_UPDATE_MODE)
                        {
                            Form.Close();
                        }
                        break;
            //            if(pVal.ItemUID == "btnFetch")
            //            {
            //                //_Form.Freeze(true);
            //                //try
            //                //{
            //                //    //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("WOD_mtrx").Specific;
            //                //    //SAPbouiCOM.DBDataSource oChildDS = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSC");
            //                //    SAPbouiCOM.EditText txtLotNo = (SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific;

            //                //    string lotNo = txtLotNo.Value.Trim();
            //                //    if (string.IsNullOrEmpty(lotNo))
            //                //    {
            //                //        Utilities.Application.SBO_Application.StatusBar.SetText("Please select a Lot No first.", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Warning);
            //                //        return;
            //                //    }

            //                //    // Get engine-chassis list from @ENGCHASISMMC for this Lot
            //                //    string queryEngChas = $@"SELECT T0.""U_EngineNo"", T0.""U_ChasisNo""
            //                //     FROM ""@ENGCHASISMMC"" T0
            //                //     INNER JOIN ""@ENGCHASISMMH"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
            //                //     WHERE T1.""U_LotNo"" = '{lotNo}'";

            //                //    SAPbobsCOM.Recordset rsEngChas = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                //    rsEngChas.DoQuery(queryEngChas);

            //                //    // Get stage count from ORST
            //                //    SAPbobsCOM.Recordset rsStage = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                //    rsStage.DoQuery(@"SELECT COUNT(""Code"") - 1 AS ""StageCount"" FROM ORST");

            //                //    int stageCount = rsStage.Fields.Item("StageCount").Value;
            //                //    int rowNo = 1;
            //                //    int currentRow = 0;

            //                //    _Matrix.Clear();
            //                //    _ChildDBDataSource.Clear();

            //                //    while (!rsEngChas.EoF)
            //                //    {
            //                //        string engineNo = rsEngChas.Fields.Item("U_EngineNo").Value.ToString();
            //                //        string chasisNo = rsEngChas.Fields.Item("U_ChasisNo").Value.ToString();

            //                //        // Repeat for each stage
            //                //        for (int i = 0; i < stageCount; i++)
            //                //        {
            //                //            _ChildDBDataSource.InsertRecord(currentRow);
            //                //            _ChildDBDataSource.Offset = currentRow;

            //                //            _ChildDBDataSource.SetValue("U_SrNo", currentRow, rowNo.ToString());
            //                //            _ChildDBDataSource.SetValue("U_EngineNo", currentRow, engineNo);
            //                //            _ChildDBDataSource.SetValue("U_ChasisNo", currentRow, chasisNo);

            //                //            rowNo++;
            //                //            currentRow++;
            //                //        }

            //                //        rsEngChas.MoveNext();
            //                //    }

            //                //    _Matrix.LoadFromDataSource();

            //                //    Utilities.Application.SBO_Application.StatusBar.SetText("Matrix filled successfully!", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Success);
            //                //}
            //                //catch (Exception ex)
            //                //{
            //                //    Utilities.Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
            //                //}
            //                //finally
            //                //{
            //                //    _Form.Freeze(false);
            //                //}
            //                _Form.Freeze(true);
            //                try
            //                {
            //                    //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("WOD_mtrx").Specific;
            //                    //SAPbouiCOM.DBDataSource oChildDS = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSC");
            //                    SAPbouiCOM.EditText txtLotNo = (SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific;
            //                    SAPbouiCOM.EditText txtPOEnt = (SAPbouiCOM.EditText)_Form.Items.Item("txtPOEnt").Specific;
            //                    SAPbouiCOM.EditText txtDocNo = (SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific;

            //                    string lotNo = txtLotNo.Value.Trim();
            //                    string prodEnt = txtPOEnt.Value.Trim();
            //                    string DocNo = txtDocNo.Value.Trim();
            //                    if (string.IsNullOrEmpty(lotNo))
            //                    {
            //                       Utilities.Application.SBO_Application.StatusBar.SetText("Please select a Lot No first.",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Warning);
            //                        return;
            //                    }

            //                    //Get engine-chassis list for this Lot
            //                    //string queryEngChas = $@"
            //                    //   SELECT T0.""U_EngineNo"", T0.""U_ChasisNo""
            //                    // FROM ""@ENGCHASISMMC"" T0
            //                    // INNER JOIN ""@ENGCHASISMMH"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
            //                    // WHERE T1.""U_LotNo"" = '{lotNo}' and T0.""U_EngineNo"" IS NOT NULL";

            //                    //string queryEngChas = $@"
            //                    //   SELECT T0.""U_EngineNo"", T0.""U_ChasisNo""
            //                    // FROM ""@ENGCHASPO"" T0
            //                    // INNER JOIN ""OWOR"" T1 ON T0.""U_ProdDocEntry"" = T1.""DocEntry""
            //                    // WHERE T1.""U_LotNo"" = '{lotNo}' and T0.""U_EngineNo"" IS NOT NULL";

            //                    string queryEngChas = $@"
            //                       SELECT T0.""U_EngineNo"", T0.""U_ChasisNo""
            //                     FROM ""@ENGCHASPO"" T0
            //                     INNER JOIN ""OWOR"" T1 ON T0.""U_ProdDocEntry"" = T1.""DocEntry""
            //                     WHERE T1.""DocEntry"" = '{prodEnt}' and T0.""U_EngineNo"" IS NOT NULL";

            //                    SAPbobsCOM.Recordset rsEngChas =(SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                    rsEngChas.DoQuery(queryEngChas);

            //                    //Get all stages (exclude overhead)
            //                    string queryStages = @"
            //                    SELECT ""Code"", ""Desc"", ""AbsEntry""
            //                    FROM ORST
            //                    WHERE ""AbsEntry"" <> '5'
            //                    ORDER BY ""AbsEntry""";

            //                    SAPbobsCOM.Recordset rsStages =(SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                    rsStages.DoQuery(queryStages);

            //                    _Matrix.Clear();
            //                    _ChildDBDataSource.Clear();

            //                    int rowNo = 1;
            //                    int currentRow = 0;
            //                    int workCounter = 1;

            //                    // Loop through each engine/chassis
            //                    while (!rsEngChas.EoF)
            //                    {
            //                        string engineNo = rsEngChas.Fields.Item("U_EngineNo").Value.ToString();
            //                        string chasisNo = rsEngChas.Fields.Item("U_ChasisNo").Value.ToString();

            //                        // Loop through each stage
            //                        rsStages.MoveFirst();
            //                        while (!rsStages.EoF)
            //                        {
            //                            _ChildDBDataSource.InsertRecord(currentRow);
            //                            _ChildDBDataSource.Offset = currentRow;
                                        
            //                            //for first barcode
            //                            string routeId = rsStages.Fields.Item("AbsEntry").Value.ToString();
            //                            string queryFBarcode = $@"SELECT T0.""ItemCode""
            //                            FROM WOR1 T0
            //                            WHERE T0.""DocEntry"" = '{productionNo}'
            //                              AND T0.""StageId"" = '{routeId}'
            //                              AND T0.""LineNum"" = (
            //                                  SELECT MIN(T1.""LineNum"")
            //                                  FROM WOR1 T1
            //                                  WHERE T1.""DocEntry"" = '{productionNo}' AND T1.""StageId"" = '{routeId}')";
            //                            SAPbobsCOM.Recordset rsFBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                            rsFBarcode.DoQuery(queryFBarcode);

            //                            //for last barcode
            //                            string queryLBarcode = $@"SELECT T0.""ItemCode""
            //                            FROM WOR1 T0
            //                            WHERE T0.""DocEntry"" = '{productionNo}'
            //                              AND T0.""StageId"" = '{routeId}'
            //                              AND T0.""LineNum"" = (
            //                                  SELECT MAX(T1.""LineNum"")
            //                                  FROM WOR1 T1
            //                                  WHERE T1.""DocEntry"" = '{productionNo}' AND T1.""StageId"" = '{routeId}')";
            //                            SAPbobsCOM.Recordset rsLBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //                            rsLBarcode.DoQuery(queryLBarcode);


            //                            _ChildDBDataSource.SetValue("U_SrNo", currentRow, rowNo.ToString());
            //                            _ChildDBDataSource.SetValue("U_EngineNo", currentRow, engineNo);
            //                            _ChildDBDataSource.SetValue("U_ChasisNo", currentRow, chasisNo);
            //                            _ChildDBDataSource.SetValue("U_RouteId", currentRow, rsStages.Fields.Item("AbsEntry").Value.ToString());
            //                            _ChildDBDataSource.SetValue("U_RouteDsp", currentRow, rsStages.Fields.Item("Desc").Value.ToString());
            //                            _ChildDBDataSource.SetValue("U_Qty", currentRow, "1");
            //                            if(!rsFBarcode.EoF)
            //                            {
            //                                _ChildDBDataSource.SetValue("U_FirstBarcode", currentRow, rsFBarcode.Fields.Item("ItemCode").Value.ToString());
            //                            }
            //                            if (!rsLBarcode.EoF)
            //                            {
            //                                _ChildDBDataSource.SetValue("U_LastBarcode", currentRow, rsLBarcode.Fields.Item("ItemCode").Value.ToString());
            //                            }

            //                            string counter = workCounter.ToString("D2");
            //                            //add engine no last 5 character to the unique workId (remove prodEnt and DocNo and only make D2)
            //                            //string workID = lotNo + "-" + prodEnt + "-" + DocNo + "-" + counter;

            //                            string chasisSuffix = chasisNo; // Default to full no if short
            //                            if (!string.IsNullOrEmpty(chasisNo) && chasisNo.Length >= 5)
            //                            {
            //                                chasisSuffix = chasisNo.Substring(chasisNo.Length - 5);
            //                            }

            //                            string workID = lotNo + "-" + chasisSuffix +  "-" + counter;

            //                            _ChildDBDataSource.SetValue("U_WorkId", currentRow, workID);

            //                            workCounter++;

            //                            rowNo++;
            //                            currentRow++;
            //                            rsStages.MoveNext();
            //                        }

            //                        //insert a blank row after different eng/chasis No
            //                        _ChildDBDataSource.InsertRecord(currentRow);
            //                        _ChildDBDataSource.Offset = currentRow;
            //                        //_ChildDBDataSource.SetValue("U_Qty", currentRow, "");
            //                        currentRow++;


            //                        rsEngChas.MoveNext();
            //                    }

            //                    _Matrix.LoadFromDataSource();

            //                    Utilities.Application.SBO_Application.StatusBar.SetText("Matrix filled successfully!",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Success);
            //                }
            //                catch (Exception ex)
            //                {
            //                    Utilities.Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message,BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Error);
            //                }
            //                finally
            //                {
            //                    _Form.Freeze(false);
            //                }
            //            }
            //            break;

            //            #endregion
            }
            }
            //#endregion
        }
        #endregion

        #region FormDefault
        internal void FormDefault()
        {
            _Form.DataBrowser.BrowseBy = "txtDocE";
            _Form.Settings.Enabled = true;

            _Form.Items.Item("txtDocE").AffectsFormMode = true;
            UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            //_Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("WOD_mtrx").Specific;
            //_MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSH");
            //_ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSC");

            //((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific).Value = (Utilities.getMaxColumnValueNum("@WRKORDRDTLSH", "U_DocNo"));
            //((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Open";
            //((SAPbouiCOM.EditText)_Form.Items.Item("txtDate").Specific).Value = DateTime.Now.ToString("yyyyMMdd");

            LoadBreakCategoryValues();
        }
        #endregion


        #region FILTER Item Type
        internal void FilterItemType(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
               //_Con.BracketOpenNum = 1;
                _Con.Alias = "U_LotNo";
                _Con.Operation = BoConditionOperation.co_NOT_NULL;
                //_Con.CondVal = ((SAPbouiCOM.EditText)_Form.Items.Item("6").Specific).Value;
                //_Con.BracketCloseNum = 1;
                _Con.Relationship = BoConditionRelationship.cr_AND;

                _Con = _Cons.Add();
                //_Con.BracketOpenNum = 1;
                _Con.Alias = "Status";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "Released";
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

        #region Production Order No
        internal void FilterProducNo(string CFLId)
        {

            _Form.Freeze(true);
            try
            {
                _Cons = new Conditions();
                _Con = _Cons.Add();
                _Con.BracketOpenNum = 1;
                _Con.Alias = "U_LotNo";
                _Con.Operation = BoConditionOperation.co_NOT_NULL;
                _Con.BracketCloseNum = 1;
                _Form.ChooseFromLists.Item("CFL_PrdONo").SetConditions(_Cons);

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

        private void SetRemarkFieldsState(string status)
        {
            try
            {
                // 1. Handle Header Remark fields (as you did in EventListener)
                _Form.Items.Item("txtRmkIn").Enabled = (status == "Start");

                bool isStop = (status == "Stop");
                //_Form.Items.Item("lblRmkOut").Visible = isStop;
                //_Form.Items.Item("txtRmkOut").Visible = isStop;
                //_Form.Items.Item("txtRmkOut").Enabled = isStop;

                // 2. Handle Matrix Column (Break Down Remarks)
                // UID from your XML: "colRemark"
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtBreak").Specific;

                // SMOOTH LOGIC:
                // If the job is Stopped, nobody should edit remarks anymore.
                // If the job is Paused, the operator MUST be able to write remarks.
                if (status == "Stop")
                {
                    oMatrix.Columns.Item("colRemark").Editable = false;
                    oMatrix.Columns.Item("colBrkCat").Editable = false;
                }
                else
                {
                    // Editable during Start, Pause, or Resume
                    oMatrix.Columns.Item("colRemark").Editable = true;
                    oMatrix.Columns.Item("colBrkCat").Editable = true;
                }
            }
            catch { }
        }


        private void LoadBreakCategoryValues()
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtBreak").Specific;
                SAPbouiCOM.Column oColumn = oMatrix.Columns.Item("colBrkCat");

                // Clear existing values first
                while (oColumn.ValidValues.Count > 0)
                {
                    oColumn.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                // Optional blank value
                oColumn.ValidValues.Add("", "");

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = @"SELECT ""Code"", ""Name""
                         FROM ""@BRKCATEGORY""
                         WHERE ""Code"" IS NOT NULL
                         ORDER BY ""Code""";

                rs.DoQuery(query);

                while (!rs.EoF)
                {
                    string code = rs.Fields.Item("Code").Value.ToString().Trim();
                    string name = rs.Fields.Item("Name").Value.ToString().Trim();

                    oColumn.ValidValues.Add(code, code);

                    rs.MoveNext();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage("Error loading Break Category values : " + ex.Message);
            }
        }
    }
}
