using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Production_Planning_LMC
{
    class ClsWorkOrderDetails : Base
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
        string docNo;
        #endregion

        #region CONSTRUCTOR and DISTRUCTOR
        public ClsWorkOrderDetails() : base()
        {

        }
        ~ClsWorkOrderDetails()
        {

        }
        #endregion

        #region Form Load
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            base.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);
            if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
            {
                this.Form.Freeze(true);
                //Utilities.SerializedMartix(ref _Form, ref _Matrix);
                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    _Form.Items.Item("txtPOEnt").Enabled = false;
                }
                this.Form.Freeze(false);
            }

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

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == BoFormMode.fm_ADD_MODE)
                        {
                            if(((SAPbouiCOM.EditText)_Form.Items.Item("txtPOEnt").Specific).Value == "")
                            {
                                Utilities.Message("Production Order No cannot be blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtInvtTr").Specific).Value == "0")
                            {
                                Utilities.Message("Work Order Details cannot be created without completing the Inventory Transfer Entry.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:

                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        //if (pVal.ItemUID == "txtLotNo")
                        //{
                        //    this.FilterItemType("");
                        //}
                        if (pVal.ItemUID == "txtPOEnt")
                        {
                            this.FilterProducNo("");
                        }


                        break;
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

                        if (pVal.ItemUID == "txtPOEnt" && _SysCFLEvent.SelectedObjects != null)
                        {
                            if (_isHandlingCFLEvent)
                            {
                                return; // Exit if we are already in this handler
                            }

                            _isHandlingCFLEvent = true;
                            try
                            {
                                _Form.Freeze(true);

                                productionNo = _SysCFLEvent.SelectedObjects.GetValue("DocEntry", 0).ToString().Trim();
                                _MasterDBDataSource.SetValue("U_PrdOrdEnt", 0, productionNo);

                                SAPbobsCOM.Recordset rsLot = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string queryLot = $@"SELECT COUNT(*) AS ""CNT"", MAX(""DocEntry"") AS ""DocEntry"" FROM ""@WRKORDRDTLSH"" WHERE ""U_PrdOrdEnt"" = '{productionNo.Replace("'", "''")}'";
                                rsLot.DoQuery(queryLot);
                                int count = Convert.ToInt32(rsLot.Fields.Item("CNT").Value);

                                if (count == 0)
                                {

                                    // Fetch Production Order Number and Product Details
                                    //string query = $@"SELECT ""DocEntry"", ""DocNum"", ""ItemCode"", ""ProdName"" FROM OWOR WHERE ""U_LotNo"" = '{selectedLotNo}'";
                                    //string query = $@"SELECT T0.""DocNum"", T0.""ItemCode"", T0.""ProdName"", T0.""PlannedQty"", T1.""U_LotNo"", T1.""U_EnChNo"", 
                                    //T0.""U_InvtTransferReq"", T5.""DocEntry"" AS ""InvtTransNo"", T0.""U_InvtTransfer""
                                    //FROM OWOR T0  
                                    //INNER JOIN ""@ENGCHASPO"" T1 ON T0.""DocEntry"" = T1.""U_ProdDocEntry"" 
                                    //LEFT JOIN ""OWTQ"" T2 ON  T0.""U_InvtTransferNo"" = T2.""DocEntry""
                                    // INNER JOIN ""WTQ1"" T3 ON T2.""DocEntry"" = T3.""DocEntry""
                                    // INNER JOIN ""WTR1"" T4 ON T4.""BaseEntry"" = T3.""DocEntry""
                                    //INNER JOIN ""OWTR"" T5 ON T4.""DocEntry"" = T5.""DocEntry""
                                    //WHERE T0.""DocEntry"" = '{productionNo}'";
                                    string query = $@"SELECT T0.""DocNum"", T0.""ItemCode"", T0.""ProdName"", T0.""PlannedQty"", T1.""U_LotNo"", T1.""U_EnChNo"", 
                                    T0.""U_InvtTransferReq"", T0.""U_InvtTransfer""
                                    FROM OWOR T0  
                                    INNER JOIN ""@ENGCHASPO"" T1 ON T0.""DocEntry"" = T1.""U_ProdDocEntry"" 
                                    WHERE T0.""DocEntry"" = '{productionNo}'";
                                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    rs.DoQuery(query);

                                    if (!rs.EoF)
                                    {
                                        //string docEntry = rs.Fields.Item("DocEntry").Value.ToString();
                                        string docNum = rs.Fields.Item("DocNum").Value.ToString();
                                        string itemCode = rs.Fields.Item("ItemCode").Value.ToString();
                                        string itemName = rs.Fields.Item("ProdName").Value.ToString();
                                        string lotNo = rs.Fields.Item("U_LotNo").Value.ToString();
                                        string plannedQty = rs.Fields.Item("PlannedQty").Value.ToString();
                                        string engineDocEntry = rs.Fields.Item("U_EnChNo").Value.ToString();
                                        string InvtTransfReqNo = rs.Fields.Item("U_InvtTransfer").Value.ToString();
                                        //_MasterDBDataSource.SetValue("U_ProdOrdEntry", 0, docEntry);
                                        _MasterDBDataSource.SetValue("U_ProdOrdNo", 0, docNum);
                                        _MasterDBDataSource.SetValue("U_ProductCode", 0, itemCode);
                                        _MasterDBDataSource.SetValue("U_ProductName", 0, itemName);
                                        //_MasterDBDataSource.SetValue("U_LotNo", 0, lotNo);
                                        _MasterDBDataSource.SetValue("U_PlannedQty", 0, plannedQty);
                                        _MasterDBDataSource.SetValue("U_InvtTransferNo", 0, InvtTransfReqNo);

                                        string modelQuery = $@"Select T0.""ItmsGrpNam"" from OITB T0 inner join OITM T1 on T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" where T1.""ItemCode"" = '{itemCode}' ";
                                        SAPbobsCOM.Recordset rsModel = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rsModel.DoQuery(modelQuery);

                                        string model = rsModel.Fields.Item("ItmsGrpNam").Value.ToString();
                                        _MasterDBDataSource.SetValue("U_Model", 0, model);

                                        FillWorkOrderMatrix(productionNo);

                                        //fetch Doc No of Engine/Chasis Mapping Master
                                        string query1 = $@"SELECT ""U_DocNo"" FROM ""@ENGCHASISMMH"" WHERE ""DocEntry"" = '{engineDocEntry}'";
                                        SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rs1.DoQuery(query1);
                                        if (!rs1.EoF)
                                        {
                                            docNo = rs1.Fields.Item("U_DocNo").Value.ToString();
                                            //_MasterDBDataSource.SetValue("U_AdvancedSBNo", 0, docNo);
                                        }
                                        else
                                        {
                                            Utilities.ShowErrorMessage("No Engine/Chasis Mapping Master found for Lot No: " + selectedLotNo);
                                        }
                                    }
                                    else
                                    {
                                        Utilities.ShowErrorMessage("No Production Order found for Lot No: " + selectedLotNo);
                                    }

                                    //// Prepare Matrix for user input
                                    //_Matrix.Clear();
                                    //_ChildDBDataSource.Clear();
                                    //_ChildDBDataSource.InsertRecord(0);

                                    //_Matrix.LoadFromDataSource();
                                    //AddSequenceNumbersToMatrix(_Matrix);
                                }
                                else if (count == 1)
                                {
                                    SAPbobsCOM.Recordset rsLot1 = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string codeQuery1 = $@"SELECT ""U_PrdOrdEnt""
                                     FROM ""@WRKORDRDTLSH"" 
                                     WHERE ""U_PrdOrdEnt"" = '{productionNo.Replace("'", "''")}'";

                                    rsLot1.DoQuery(codeQuery1);
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                                    string recordLotNo = rsLot1.Fields.Item("U_PrdOrdEnt").Value.ToString();

                                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtPOEnt").Specific).Value = recordLotNo;
                                    _MasterDBDataSource.SetValue("U_PrdOrdEnt", 0, recordLotNo);

                                    System.Threading.Thread.Sleep(60);

                                    _Form.Items.Item("1").Click();

                                    System.Threading.Thread.Sleep(60);
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.ShowErrorMessage("Error while setting LotNo CFL value: " + ex.Message);
                            }
                            finally
                            {
                                _Form.Freeze(false);

                                _isHandlingCFLEvent = false;
                            }
                            
                        }

                        break;

                    #endregion

                    #region Item Pressed (button Fetch Details)
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if(pVal.ItemUID == "btnFetch")
                        //{
                        //    _Form.Freeze(true);
                        //    try
                        //    {
                        //        //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("WOD_mtrx").Specific;
                        //        //SAPbouiCOM.DBDataSource oChildDS = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSC");
                        //        //SAPbouiCOM.EditText txtLotNo = (SAPbouiCOM.EditText)_Form.Items.Item("txtLotNo").Specific;
                        //        SAPbouiCOM.EditText txtPOEnt = (SAPbouiCOM.EditText)_Form.Items.Item("txtPOEnt").Specific;
                        //        SAPbouiCOM.EditText txtDocNo = (SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific;

                        //        //string lotNo = txtLotNo.Value.Trim();
                        //        string prodEnt = txtPOEnt.Value.Trim();
                        //        string DocNo = txtDocNo.Value.Trim();
                        //        //if (string.IsNullOrEmpty(lotNo))
                        //        //{
                        //        //   Utilities.Application.SBO_Application.StatusBar.SetText("Please select a Lot No first.",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Warning);
                        //        //    return;
                        //        //}

                        //        string queryEngChas = $@"
                        //           SELECT T0.""U_EngineNo"", T0.""U_ChasisNo"", T0.""U_LotNo"", T0.""U_EnChNo""
                        //         FROM ""@ENGCHASPO"" T0
                        //         INNER JOIN ""OWOR"" T1 ON T0.""U_ProdDocEntry"" = T1.""DocEntry""
                        //         WHERE T1.""DocEntry"" = '{prodEnt}' and T0.""U_EngineNo"" IS NOT NULL";

                        //        SAPbobsCOM.Recordset rsEngChas =(SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //        rsEngChas.DoQuery(queryEngChas);

                        //        //Get all stages (exclude overhead)
                        //        string queryStages = @"
                        //        SELECT ""Code"", ""Desc"", ""AbsEntry""
                        //        FROM ORST
                        //        WHERE ""AbsEntry"" <> '5'
                        //        ORDER BY ""AbsEntry""";

                        //        SAPbobsCOM.Recordset rsStages =(SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //        rsStages.DoQuery(queryStages);

                        //        //get the stages for each production 
                        //        string queryStagesProd = $@"
                        //        SELECT distinct T0.""StageId"", T1.""Desc""
                        //        FROM WOR1 T0
                        //        INNER JOIN ORST T1 ON T0.""StageId"" = T1.""AbsEntry""
                        //        WHERE ""DocEntry"" = '{prodEnt}' and T0.""StageId"" <> '5'
                        //        ORDER BY T0.""StageId"" ";
                        //        SAPbobsCOM.Recordset rsStagesProd = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //        rsStagesProd.DoQuery(queryStagesProd);

                        //        _Matrix.Clear();
                        //        _ChildDBDataSource.Clear();

                        //        int rowNo = 1;
                        //        int currentRow = 0;
                        //        int workCounter = 1;

                        //        // Loop through each engine/chassis
                        //        while (!rsEngChas.EoF)
                        //        {
                        //            string engineNo = rsEngChas.Fields.Item("U_EngineNo").Value.ToString();
                        //            string chasisNo = rsEngChas.Fields.Item("U_ChasisNo").Value.ToString();
                        //            string lot = rsEngChas.Fields.Item("U_LotNo").Value.ToString();
                        //            string EngineChasDocEntry = rsEngChas.Fields.Item("U_EnChNo").Value.ToString();

                        //            // Loop through each stage
                        //            rsStagesProd.MoveFirst();
                        //            while (!rsStagesProd.EoF)
                        //            {
                        //                _ChildDBDataSource.InsertRecord(currentRow);
                        //                _ChildDBDataSource.Offset = currentRow;

                        //                //for first barcode
                        //                //string queryFBarcode = $@"SELECT T0.""ItemCode""
                        //                //FROM WOR1 T0
                        //                //WHERE T0.""DocEntry"" = '{productionNo}'
                        //                //  AND T0.""StageId"" = '{routeId}'
                        //                //  AND T0.""LineNum"" = (
                        //                //      SELECT MIN(T1.""LineNum"")
                        //                //      FROM WOR1 T1
                        //                //      WHERE T1.""DocEntry"" = '{productionNo}' AND T1.""StageId"" = '{routeId}')";

                        //                //for last barcode
                        //                //string queryLBarcode = $@"SELECT T0.""ItemCode""
                        //                //FROM WOR1 T0
                        //                //WHERE T0.""DocEntry"" = '{productionNo}'
                        //                //  AND T0.""StageId"" = '{routeId}'
                        //                //  AND T0.""LineNum"" = (
                        //                //      SELECT MAX(T1.""LineNum"")
                        //                //      FROM WOR1 T1
                        //                //      WHERE T1.""DocEntry"" = '{productionNo}' AND T1.""StageId"" = '{routeId}')";
                        //                //SAPbobsCOM.Recordset rsLBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //                //rsLBarcode.DoQuery(queryLBarcode);

                        //                //string routeId = rsStages.Fields.Item("AbsEntry").Value.ToString();
                        //                string routeId = rsStagesProd.Fields.Item("StageId").Value.ToString();

                        //                string queryFBarcode = $@"Select ""U_JobStartBarCode"", ""U_JobStopBarCode"", ""U_JobPauseBarCode"", ""U_JobResumeBarCode"" from ORST where ""AbsEntry"" = '{routeId}'";
                        //                SAPbobsCOM.Recordset rsFBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //                rsFBarcode.DoQuery(queryFBarcode);

                        //                _ChildDBDataSource.SetValue("U_SrNo", currentRow, rowNo.ToString());
                        //                _ChildDBDataSource.SetValue("U_EngineNo", currentRow, engineNo);
                        //                _ChildDBDataSource.SetValue("U_ChasisNo", currentRow, chasisNo);
                        //                string routeDsp = rsStagesProd.Fields.Item("Desc").Value.ToString();
                        //                string jobId = routeId + " - " + routeDsp;
                        //                _ChildDBDataSource.SetValue("U_RouteId", currentRow, jobId);
                        //                _ChildDBDataSource.SetValue("U_RouteIdNum", currentRow, routeId);
                        //                _ChildDBDataSource.SetValue("U_RouteDsp", currentRow, rsStagesProd.Fields.Item("Desc").Value.ToString());
                        //                _ChildDBDataSource.SetValue("U_BatchNo", currentRow, lot);
                        //                _ChildDBDataSource.SetValue("U_Status", currentRow, "Pending");
                        //                //_ChildDBDataSource.SetValue("U_Qty", currentRow, "1");
                        //                if (!rsFBarcode.EoF)
                        //                {
                        //                    _ChildDBDataSource.SetValue("U_FirstBarcode", currentRow, rsFBarcode.Fields.Item("U_JobStartBarCode").Value.ToString());
                        //                    _ChildDBDataSource.SetValue("U_LastBarcode", currentRow, rsFBarcode.Fields.Item("U_JobStopBarCode").Value.ToString());
                        //                    _ChildDBDataSource.SetValue("U_JobPause", currentRow, rsFBarcode.Fields.Item("U_JobPauseBarCode").Value.ToString());
                        //                    _ChildDBDataSource.SetValue("U_JobResume", currentRow, rsFBarcode.Fields.Item("U_JobResumeBarCode").Value.ToString());
                        //                }
                        //                //if (!rsLBarcode.EoF)
                        //                //{
                        //                //    _ChildDBDataSource.SetValue("U_LastBarcode", currentRow, rsLBarcode.Fields.Item("ItemCode").Value.ToString());
                        //                //}

                        //                string counter = workCounter.ToString("D2");
                        //                //string workID = lotNo + "-" + prodEnt + "-" + DocNo + "-" + counter;

                        //                string chasisSuffix = chasisNo; // Default to full no if short
                        //                if (!string.IsNullOrEmpty(chasisNo) && chasisNo.Length >= 5)
                        //                {
                        //                    chasisSuffix = chasisNo.Substring(chasisNo.Length - 5);
                        //                }

                        //                string workID = lot + "-" + chasisSuffix +  "-" + counter;

                        //                _ChildDBDataSource.SetValue("U_WorkId", currentRow, workID);
                        //                _ChildDBDataSource.SetValue("U_AdvancedSBNo", currentRow, EngineChasDocEntry);

                        //                workCounter++;

                        //                rowNo++;
                        //                currentRow++;
                        //                rsStagesProd.MoveNext();
                        //            }

                        //            //insert a blank row after difference in engine/chasis No for seperation
                        //            _ChildDBDataSource.InsertRecord(currentRow);
                        //            _ChildDBDataSource.Offset = currentRow;
                        //            //_ChildDBDataSource.SetValue("U_Qty", currentRow, "");
                        //            currentRow++;
                        //            rsEngChas.MoveNext();
                        //        }

                        //        _Matrix.LoadFromDataSource();

                        //        Utilities.Application.SBO_Application.StatusBar.SetText("Data fetched successfully!",BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Success);
                        //    }
                        //    catch (Exception ex)
                        //    {
                        //        Utilities.Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message,BoMessageTime.bmt_Short,BoStatusBarMessageType.smt_Error);
                        //    }
                        //    finally
                        //    {
                        //        _Form.Freeze(false);
                        //    }
                        //}
                        break;

                        #endregion
                }
            }
            #endregion
        }
        #endregion

        #region FormDefault
        internal void FormDefault()
        {
            _Form.Freeze(true);
            _Form.DataBrowser.BrowseBy = "txtPOEnt";
            _Form.Settings.Enabled = true;

            _Form.Items.Item("txtPOEnt").AffectsFormMode = true;
            UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);

            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("WOD_mtrx").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@WRKORDRDTLSC");

            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNo").Specific).Value = (Utilities.getMaxColumnValueNum("@WRKORDRDTLSH", "U_DocNo"));
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Open";
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDate").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtUpdtBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
            string currentUser = ((SAPbouiCOM.EditText)_Form.Items.Item("txtUpdtBy").Specific).Value.Trim();
            string query = $@"select ""USER_CODE"" from OUSR where ""U_NAME"" = '{currentUser}' ";
            SAPbobsCOM.Recordset rsUserCode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rsUserCode.DoQuery(query);
            string userCode = rsUserCode.Fields.Item("USER_CODE").Value.ToString();
            _Form.Freeze(false);

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
                _Con.Alias = "Status";
                _Con.Operation = BoConditionOperation.co_EQUAL;
                _Con.CondVal = "R";
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

        private void FillWorkOrderMatrix(string prodEnt)
        {
            try
            {
                // 1. Get Engine/Chassis Data
                string queryEngChas = $@"
            SELECT T0.""U_EngineNo"", T0.""U_ChasisNo"", T0.""U_LotNo"", T0.""U_EnChNo""
            FROM ""@ENGCHASPO"" T0
            WHERE T0.""U_ProdDocEntry"" = '{prodEnt}' AND T0.""U_EngineNo"" IS NOT NULL";

                SAPbobsCOM.Recordset rsEngChas = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsEngChas.DoQuery(queryEngChas);

                // 2. Get Route Stages for this Production Order
                string queryStagesProd = $@"SELECT ""StgEntry"", ""Name"" 
                            FROM WOR4  
                            WHERE ""DocEntry"" = {prodEnt} and ""StgEntry"" <> 5 
                            ORDER BY  ""StgEntry"" ";

                SAPbobsCOM.Recordset rsStagesProd = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsStagesProd.DoQuery(queryStagesProd);

                _Matrix.Clear();
                _ChildDBDataSource.Clear();

                int rowNo = 1;
                int currentRow = 0;

                while (!rsEngChas.EoF)
                {
                    string engineNo = rsEngChas.Fields.Item("U_EngineNo").Value.ToString();
                    string chasisNo = rsEngChas.Fields.Item("U_ChasisNo").Value.ToString();
                    string lot = rsEngChas.Fields.Item("U_LotNo").Value.ToString();
                    string EngineChasDocEntry = rsEngChas.Fields.Item("U_EnChNo").Value.ToString();
                    int workCounter = 1;

                    rsStagesProd.MoveFirst();
                    while (!rsStagesProd.EoF)
                    {
                        _ChildDBDataSource.InsertRecord(currentRow);

                        string routeId = rsStagesProd.Fields.Item("StageId").Value.ToString();
                        string routeDsp = rsStagesProd.Fields.Item("Desc").Value.ToString();

                        // Get Barcodes from OSRT
                        string queryBarcodes = $@"Select ""U_JobStartBarCode"", ""U_JobStopBarCode"", ""U_JobPauseBarCode"", ""U_JobResumeBarCode"" from ORST where ""AbsEntry"" = '{routeId}'";
                        SAPbobsCOM.Recordset rsBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsBarcode.DoQuery(queryBarcodes);

                        _ChildDBDataSource.SetValue("U_SrNo", currentRow, rowNo.ToString());
                        _ChildDBDataSource.SetValue("U_EngineNo", currentRow, engineNo);
                        _ChildDBDataSource.SetValue("U_ChasisNo", currentRow, chasisNo);
                        _ChildDBDataSource.SetValue("U_RouteId", currentRow, routeId + " - " + routeDsp);
                        _ChildDBDataSource.SetValue("U_RouteIdNum", currentRow, routeId);
                        _ChildDBDataSource.SetValue("U_RouteDsp", currentRow, routeDsp);
                        _ChildDBDataSource.SetValue("U_BatchNo", currentRow, lot);
                        _ChildDBDataSource.SetValue("U_Status", currentRow, "Pending");

                        if (!rsBarcode.EoF)
                        {
                            _ChildDBDataSource.SetValue("U_FirstBarcode", currentRow, rsBarcode.Fields.Item("U_JobStartBarCode").Value.ToString());
                            _ChildDBDataSource.SetValue("U_LastBarcode", currentRow, rsBarcode.Fields.Item("U_JobStopBarCode").Value.ToString());
                            _ChildDBDataSource.SetValue("U_JobPause", currentRow, rsBarcode.Fields.Item("U_JobPauseBarCode").Value.ToString());
                            _ChildDBDataSource.SetValue("U_JobResume", currentRow, rsBarcode.Fields.Item("U_JobResumeBarCode").Value.ToString());
                        }

                        // WorkID Generation
                        string chasisSuffix = chasisNo.Length >= 5 ? chasisNo.Substring(chasisNo.Length - 5) : chasisNo;
                        string workID = $"{lot}-{chasisSuffix}-{workCounter.ToString("D2")}";

                        _ChildDBDataSource.SetValue("U_WorkId", currentRow, workID);
                        _ChildDBDataSource.SetValue("U_AdvancedSBNo", currentRow, EngineChasDocEntry);

                        workCounter++;
                        rowNo++;
                        currentRow++;
                        rsStagesProd.MoveNext();
                    }

                    // Insert a blank row for separation
                    _ChildDBDataSource.InsertRecord(currentRow);
                    currentRow++;
                    rsEngChas.MoveNext();
                }

                _Matrix.LoadFromDataSource();
            }
            catch (Exception ex)
            {
                throw new Exception("Error Filling Matrix: " + ex.Message);
            }
        }
    }
}
