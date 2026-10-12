using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Xml;
using System.Linq;
using System.Threading.Tasks;

namespace Production_Planning_LMC
{
    class ClsProductionOrderr : Base
    {
        #region CONSTRUCTOR
        public ClsProductionOrderr() : base() { }
        ~ClsProductionOrderr() { }
        #endregion

        #region VARIABLE DECLARATION
        string lotNo;
        int engChasMMDocEntry;
        string itemCode, status;
        int selectedRowIndex;
        private System.Collections.Generic.HashSet<int> frozenRows = new System.Collections.Generic.HashSet<int>();
        SAPbobsCOM.Recordset rs1 = null;
        int TransferReqDocEntry, TransferDocEntry;
        bool hasError = false;
        string errorMsg = "";
        private string lastProcessedItem = "";
        #endregion

        #region getter and setter for Form
        public SAPbouiCOM.Form Form { get; set; }
        #endregion

        #region FormDefault
        public void FormDefault()
        {
            try
            {
                if (ItemExists(this.Form, "lblVer")) return;

                SAPbouiCOM.Item oItemStatus = this.Form.Items.Item("7");
                SAPbouiCOM.Item oItemStatusLabel = this.Form.Items.Item("8"); 

                // Label
                SAPbouiCOM.Item oLabelItem = this.Form.Items.Add("lblOCN", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oLabelItem.Left = oItemStatusLabel.Left + 300;
                oLabelItem.Top = oItemStatusLabel.Top + oItemStatusLabel.Height ; // Positioned right below 'Status' label
                oLabelItem.Width = oItemStatusLabel.Width;
                oLabelItem.Height = oItemStatusLabel.Height;
                oLabelItem.FromPane = 0; // Visible on all tabs
                oLabelItem.ToPane = 0;

                SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oLabelItem.Specific;
                oStaticText.Caption = "OCN Code";

                //Input Field
                SAPbouiCOM.Item oEditItem = this.Form.Items.Add("txtOCN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oEditItem.Left = oItemStatus.Left + 500;
                oEditItem.Top = oItemStatus.Top + oItemStatus.Height; // Positioned right below 'Status' combo box
                oEditItem.Width = oItemStatus.Width;
                oEditItem.Height = oItemStatus.Height;
                oEditItem.FromPane = 0;
                oEditItem.ToPane = 0;

                //SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oEditItem.Specific;

                ////bind
                //oEditText.DataBind.SetBound(true, "OWOR", "U_OCNCode");
                //oEditItem.Enabled = false;

                // 1. Create a memory field (UserDataSource)
                this.Form.DataSources.UserDataSources.Add("udsOCN", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);

                // 2. Bind the UI EditText to the UserDataSource
                SAPbouiCOM.Item oItem = this.Form.Items.Item("txtOCN");
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                oEdit.DataBind.SetBound(true, "", "udsOCN");
                oEditItem.Enabled = false;

                if (ItemExists(this.Form, "fldEngChas"))
                {
                    this.Form.PaneLevel = 1;
                    return;
                }

                this.Form.Freeze(true);

                // created in-memory data holders instead of linking directly to the database.
                this.Form.DataSources.UserDataSources.Add("colSrNo", BoDataType.dt_SHORT_TEXT, 50);
                this.Form.DataSources.UserDataSources.Add("colCha", BoDataType.dt_SHORT_TEXT, 150);
                this.Form.DataSources.UserDataSources.Add("colEng", BoDataType.dt_SHORT_TEXT, 150);
                this.Form.DataSources.UserDataSources.Add("colSet", BoDataType.dt_SHORT_TEXT, 150);
                this.Form.DataSources.UserDataSources.Add("colTrans", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colLotNo", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colEgChNo", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colPinCod", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colStatus", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colPInvt", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colRProd", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colTTime", BoDataType.dt_SHORT_TEXT, 180);
                this.Form.DataSources.UserDataSources.Add("colBatch", BoDataType.dt_SHORT_TEXT, 100);
                this.Form.DataSources.UserDataSources.Add("colProdS", BoDataType.dt_SHORT_TEXT, 100);
                this.Form.DataSources.UserDataSources.Add("colDel", BoDataType.dt_SHORT_TEXT, 20);


                // Tab Creation 
                Item lastTab = this.Form.Items.Item("234000008");
                Item newTab = this.Form.Items.Add("fldEngChas", BoFormItemTypes.it_FOLDER);
                newTab.Left = lastTab.Left + lastTab.Width; 
                newTab.Top = lastTab.Top;
                newTab.Width = lastTab.Width + 20; 
                newTab.Height = lastTab.Height;
                Folder f = (Folder)newTab.Specific;
                f.Caption = "Engine/Chassis Selection"; f.Pane = 50; f.GroupWith("234000008");

                if (this.Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    newTab.Enabled = false;
                }

                // CFL Creation 
                ChooseFromListCollection oCFLs = this.Form.ChooseFromLists;
                ChooseFromListCreationParams oCFLParams;
                oCFLParams = (ChooseFromListCreationParams)Utilities.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
                oCFLParams.ObjectType = "10000045"; oCFLParams.UniqueID = "cflEng"; oCFLs.Add(oCFLParams);
                //oCFLParams = (ChooseFromListCreationParams)Utilities.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_ChooseFromListCreationParams);
                //oCFLParams.ObjectType = "10000045"; oCFLParams.UniqueID = "cflCha"; oCFLs.Add(oCFLParams);
  
                //Matrix Creation 
                Item oMatrixItem = this.Form.Items.Add("mtxEngChas", BoFormItemTypes.it_MATRIX);
                oMatrixItem.FromPane = 50; 
                oMatrixItem.ToPane = 50; 
                oMatrixItem.Left = 25; 
                oMatrixItem.Top = lastTab.Top + 40;
                //instead of .ClientHeight and .ClientWidth
                //oMatrixItem.Width = this.Form.Width - 290; 
                //oMatrixItem.Height = this.Form.Height - 250;
                oMatrixItem.Width = 810;
                oMatrixItem.Height = 160;
                Matrix m = (Matrix)oMatrixItem.Specific;

                //Button Creation for Inventor Transfer Request
                Item oBtnItem = this.Form.Items.Add("btnCustom", BoFormItemTypes.it_BUTTON);
                oBtnItem.FromPane = 50;
                oBtnItem.ToPane = 50;

                oBtnItem.Left = 830;
                oBtnItem.Top = oMatrixItem.Top + oMatrixItem.Height - 100;

                oBtnItem.Width = 150;
                oBtnItem.Height = 25;
                SAPbouiCOM.Button oBtn = (SAPbouiCOM.Button)oBtnItem.Specific;
                oBtn.Caption = "Inventory Transfer Request";
                oBtnItem.Enabled = false;

                //Button Creation for Inventor Transfer
                Item oBtnItemT = this.Form.Items.Add("btnInvtT", BoFormItemTypes.it_BUTTON);
                oBtnItemT.FromPane = 50;
                oBtnItemT.ToPane = 50;

                oBtnItemT.Left = 830;
                oBtnItemT.Top = oMatrixItem.Top + oMatrixItem.Height - 50;

                oBtnItemT.Width = 150;
                oBtnItemT.Height = 25;
                SAPbouiCOM.Button oBtnT = (SAPbouiCOM.Button)oBtnItemT.Specific;
                oBtnT.Caption = "Inventory Transfer";
                oBtnItemT.Enabled = false;

                // BINDING MATRIX COLUMNS TO USERDATASOURCES 
                // The table name is "" (empty string), and the alias is the UserDataSource name.
                m.Columns.Add("colSrNo", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Sr No";
                m.Columns.Item("colSrNo").DataBind.SetBound(true, "", "colSrNo");


                m.Columns.Add("colCha", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Chassis No";
                m.Columns.Item("colCha").DataBind.SetBound(true, "", "colCha");

                m.Columns.Add("colEng", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Engine No";
                m.Columns.Item("colEng").DataBind.SetBound(true, "", "colEng");

                m.Columns.Add("colSet", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Key Set";
                m.Columns.Item("colSet").DataBind.SetBound(true, "", "colSet");

                m.Columns.Add("colTrans", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Transmission No";
                m.Columns.Item("colTrans").DataBind.SetBound(true, "", "colTrans");

                m.Columns.Add("colLotNo", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Lot No";
                m.Columns.Item("colLotNo").DataBind.SetBound(true, "", "colLotNo");


                //SAPbouiCOM.Column oColLink = m.Columns.Add("colEgLink", BoFormItemTypes.it_LINKED_BUTTON);
                //oColLink.Width = 20;
                //oColLink.TitleObject.Caption = "";
                //oColLink.Editable = true;

                //oColLink.DataBind.SetBound(true, "", "colEgChNo");

                //SAPbouiCOM.LinkedButton oLink = (SAPbouiCOM.LinkedButton)oColLink.ExtendedObject;

                // DocEntry Column (Engine/Chasis Mapping Master)
                SAPbouiCOM.Column oColData = m.Columns.Add("colEgChNo", BoFormItemTypes.it_LINKED_BUTTON);
                oColData.TitleObject.Caption = "Engine/Chassis DocEntry";
                oColData.Width = 100;
                oColData.Editable = false; 
                oColData.DataBind.SetBound(true, "", "colEgChNo");
                SAPbouiCOM.LinkedButton oLinkk = (SAPbouiCOM.LinkedButton)oColData.ExtendedObject;
                oLinkk.LinkedObjectType = "ENGCHASISMM";

                m.Columns.Add("colPinCod", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Pin Code";
                m.Columns.Item("colPinCod").DataBind.SetBound(true, "", "colPinCod");

                m.Columns.Add("colStatus", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Status";
                m.Columns.Item("colStatus").DataBind.SetBound(true, "", "colStatus");

                m.Columns.Add("colPInvt", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Push To Inventory";
                m.Columns.Item("colPInvt").DataBind.SetBound(true, "", "colPInvt");

                SAPbouiCOM.Column oColDataa = m.Columns.Add("colRProd", BoFormItemTypes.it_LINKED_BUTTON);
                oColDataa.TitleObject.Caption = "Receipt from Production";
                oColDataa.DataBind.SetBound(true, "", "colRProd");
                SAPbouiCOM.LinkedButton oLinkkk = (SAPbouiCOM.LinkedButton)oColDataa.ExtendedObject;
                oLinkkk.LinkedObjectType = "59";

                m.Columns.Add("colTTime", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Total Time Consumed";
                m.Columns.Item("colTTime").DataBind.SetBound(true, "", "colTTime");

                m.Columns.Add("colBatch", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Batch No";
                m.Columns.Item("colBatch").DataBind.SetBound(true, "", "colBatch");

                m.Columns.Add("colProdS", BoFormItemTypes.it_EDIT).TitleObject.Caption = "Production Series No";
                m.Columns.Item("colProdS").DataBind.SetBound(true, "", "colProdS");

                // --- Assign CFLs & other Properties
                m.Columns.Item("colCha").ChooseFromListUID = "cflEng"; m.Columns.Item("colCha").ChooseFromListAlias = "DistNumber";
                //m.Columns.Item("colCha").ChooseFromListUID = "cflCha"; m.Columns.Item("colCha").ChooseFromListAlias = "DistNumber";

                m.Columns.Item("colSrNo").Width = 50; m.Columns.Item("colSrNo").Editable = false;
                m.Columns.Item("colEng").Width = 150; m.Columns.Item("colEng").Editable = false;
                m.Columns.Item("colCha").Width = 150; m.Columns.Item("colCha").Editable = true;
                m.Columns.Item("colSet").Width = 150; m.Columns.Item("colSet").Editable = false;
                m.Columns.Item("colTrans").Width = 180; m.Columns.Item("colTrans").Editable = false;
                m.Columns.Item("colLotNo").Width = 180; m.Columns.Item("colLotNo").Editable = false;
                m.Columns.Item("colPinCod").Width = 180; m.Columns.Item("colPinCod").Editable = false;
                m.Columns.Item("colStatus").Width = 180; m.Columns.Item("colStatus").Editable = false;
                m.Columns.Item("colPInvt").Width = 120; m.Columns.Item("colPInvt").Editable = false;
                m.Columns.Item("colRProd").Width = 120; m.Columns.Item("colRProd").Editable = false;
                m.Columns.Item("colTTime").Width = 120; m.Columns.Item("colTTime").Editable = false;
                m.Columns.Item("colBatch").Width = 120; m.Columns.Item("colBatch").Editable = false;
                m.Columns.Item("colProdS").Width = 180; m.Columns.Item("colProdS").Editable = false;


                SAPbouiCOM.Column oColDel = m.Columns.Add("colDel", BoFormItemTypes.it_EDIT);
                oColDel.TitleObject.Caption = "Action";
                oColDel.Width = 60;
                oColDel.Editable = false; 
                oColDel.DataBind.SetBound(true, "", "colDel");
                oColDel.FontSize = 12;
                oColDel.ForeColor = 255;


                m.Clear();
                m.AddRow();
              
                //when status is "Canceled" then donot load the matrix
                string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                if (loadedStatus != "C")
                {
                    LoadMatrixData();
                }
                if (loadedStatus == "R")
                {
                    this.Form.Mode = BoFormMode.fm_OK_MODE;
                }
                this.Form.PaneLevel = 1;
            }
            catch (Exception ex) 
            { 
                Utilities.Application.SBO_Application.MessageBox("FormDefault Error: " + ex.Message); 
            }
            finally 
            {
                this.Form.Items.Item("U_InvtTransfer").Enabled = false;
                this.Form.Items.Item("U_InvtTransferReq").Enabled = false;
                this.Form.Items.Item("U_WorkOrderEntry").Enabled = false;
                this.Form.Items.Item("txtOCN").Enabled = false;
                this.Form.Freeze(false); 
            }
        }
        #endregion

        #region Item Exists
        private bool ItemExists(SAPbouiCOM.Form form, string itemUID)
        {
            try { form.Items.Item(itemUID); return true; } catch { return false; }
        }
        #endregion

        #region Item Event
        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            #region BeforeAction
            if (pVal.BeforeAction)
            {
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.ItemUID == "fldEngChas")
                    {
                        try
                        {
                            this.Form.Freeze(true);
                            Folder f = (Folder)this.Form.Items.Item(pVal.ItemUID).Specific;
                            if (this.Form.PaneLevel != f.Pane)
                            {
                                this.Form.PaneLevel = f.Pane;
                            }
                        }
                        catch (Exception) { /* Handle error */ }
                        finally 
                        {
                            this.Form.Items.Item("U_InvtTransfer").Enabled = false;
                            this.Form.Items.Item("U_InvtTransferReq").Enabled = false;
                            this.Form.Items.Item("U_WorkOrderEntry").Enabled = false;
                            this.Form.Items.Item("txtOCN").Enabled = false;

                            this.Form.Freeze(false);
                        }
                    }
                }

                #region Click Event on Delete
                if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "mtxEngChas" && pVal.ColUID == "colDel" && pVal.Row > 0)
                {
                    Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                    string cellValue = ((EditText)oMatrix.Columns.Item("colDel").Cells.Item(pVal.Row).Specific).Value.Trim();
                    string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                    if (loadedStatus == "R")
                    {
                        BubbleEvent = false;
                    }
                    else
                    {
                        // Only react if the text is "Delete"
                        if (cellValue == "Delete")
                        {
                            int res = Utilities.Application.SBO_Application.MessageBox("Are you sure you want to remove this row?", 1, "Yes", "No");
                            if (res == 1)
                            {
                                DeleteSpecificRow(pVal.Row);
                                BubbleEvent = false; // Stop further processing since row is gone
                            }
                            else
                            {
                                // User Cancelled, just let event pass
                            }
                        }
                    }
                }
                #endregion

                if ((pVal.EventType == BoEventTypes.et_CLICK || pVal.EventType == BoEventTypes.et_GOT_FOCUS) && pVal.ItemUID == "mtxEngChas")
                {
                    // Check if the row the user is trying to interact with is in our frozen list
                    if (pVal.Row > 0 && frozenRows.Contains(pVal.Row))
                    {
                        // Cancel the event. This prevents the cell from getting focus.
                        BubbleEvent = false;
                        return; 
                    }
                }

                #region CFL for Chassis
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "mtxEngChas")
                {
                    string dbFieldToQuery = "";
                    switch (pVal.ColUID)
                    {
                        //case "colEng": dbFieldToQuery = "U_EngineNo"; break;
                        case "colCha": dbFieldToQuery = "U_ChasisNo"; break;
                    }

                    if (!string.IsNullOrEmpty(dbFieldToQuery))
                    {
                        try
                        {

                            Matrix oMatrix = (Matrix)this.Form.Items.Item(pVal.ItemUID).Specific;
                            string cflUID = oMatrix.Columns.Item(pVal.ColUID).ChooseFromListUID;
                            SAPbouiCOM.ChooseFromList oCFL = this.Form.ChooseFromLists.Item(cflUID);

                            oCFL.SetConditions(null); 

                            // fetched Lot Number from the Production Order                                         
                            //lotNo = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_LotNo", 0).Trim();
                            itemCode = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("ItemCode", 0).Trim();

                            //if (!string.IsNullOrEmpty(lotNo))
                            //{

                            // Collect all values from the current column that are already used in other rows.
                            var valuesToExclude = new System.Collections.Generic.List<string>();
                            for (int i = 1; i <= oMatrix.RowCount; i++)
                            {
                                if (i == pVal.Row) continue; // Skip the row the user is currently on.

                                string cellValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(i).Specific).Value.Trim();
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    valuesToExclude.Add(cellValue);
                                }
                            }

                            //  used "NOT IN" if the values need to be excluded in the matrix next row
                            string exclusionClause = "";
                            if (valuesToExclude.Count > 0)
                            {
                                var sanitizedValues = valuesToExclude.Select(v => $"'{v.Replace("'", "''")}'");
                                exclusionClause = $" AND T2.\"{dbFieldToQuery}\" NOT IN ({string.Join(",", sanitizedValues)})";
                            }


                            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                                //string query = $"SELECT T2.\"DocEntry\", T2.\"{dbFieldToQuery}\", T2.\"U_ModelCode\", T1.\"U_DocNo\" FROM \"@ENGCHASISMMH\" T1 ";
                                //query += "INNER JOIN \"@ENGCHASISMMC\" T2 ON T1.\"DocEntry\" = T2.\"DocEntry\" ";
                                //query += $"WHERE T2.\"{dbFieldToQuery}\" IS NOT NULL AND  T2.\"U_ModelCode\" = '{itemCode}' AND T2.\"U_ProdOrdNo\" IS NULL ";

                                string query = $"SELECT T2.\"DocEntry\", T2.\"U_ChasisNo\", T2.\"U_ModelCode\", T1.\"U_DocNo\" FROM \"@ENGCHASISMMH\" T1 ";
                                query += "INNER JOIN \"@ENGCHASISMMC\" T2 ON T1.\"DocEntry\" = T2.\"DocEntry\" ";
                                query += $"WHERE T2.\"U_ChasisNo\" IS NOT NULL AND  T2.\"U_ModelCode\" = '{itemCode}' AND T2.\"U_ProdOrdNo\" IS NULL ";
                                query += exclusionClause;

                                rs.DoQuery(query);

                                if (rs.RecordCount > 0)
                                {
                                    engChasMMDocEntry = (int)rs.Fields.Item("DocEntry").Value;

                                    Conditions oCons = (Conditions)Utilities.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_Conditions);
                                    Condition oCon;

                                    for (int i = 0; i < rs.RecordCount; i++)
                                    {
                                        oCon = oCons.Add();
                                        oCon.Alias = "DistNumber";
                                        oCon.Operation = BoConditionOperation.co_EQUAL;
                                        oCon.CondVal = rs.Fields.Item(1).Value.ToString();
                                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                                        oCon = oCons.Add();
                                        oCon.Alias = "Quantity";
                                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                        oCon.CondVal = "1";
                                        oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                                        oCon = oCons.Add();
                                        oCon.Alias = "QuantOut";
                                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                        oCon.CondVal = "0";

                                    if (i < rs.RecordCount - 1)
                                        {
                                            oCon.Relationship = BoConditionRelationship.cr_OR;
                                        }
                                        rs.MoveNext();
                                    }
                                    oCFL.SetConditions(oCons);
                                }
                                else
                                {
                                    Conditions oCons = (Conditions)Utilities.Application.SBO_Application.CreateObject(BoCreatableObjectType.cot_Conditions);
                                    Condition oCon = oCons.Add();
                                    oCon.Alias = "DistNumber";
                                    oCon.Operation = BoConditionOperation.co_EQUAL;
                                    oCon.CondVal = System.Guid.NewGuid().ToString();
                                    oCFL.SetConditions(oCons);
                                }
                            //}
                        }
                        catch (Exception ex)
                        {
                            Utilities.Application.SBO_Application.MessageBox("Error during CFL filtering: " + ex.Message);
                        }
                    }
                }
                #endregion

                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "12")
                {
                    try
                    {
                        string newValStr = ((SAPbouiCOM.EditText)this.Form.Items.Item("12").Specific).Value;

                        if (double.TryParse(newValStr, out double newPlannedQty))
                        {
                            if (newPlannedQty % 1 != 0)
                            {
                                Utilities.Application.SBO_Application.SetStatusBarMessage(
                                    "Error: Planned Quantity cannot be a decimal value. Please enter a whole number.",
                                    BoMessageTime.bmt_Short, true);

                                BubbleEvent = false;
                                return;
                            }

                            string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                            Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                            int currentMatrixRows = oMatrix.VisualRowCount;

                            // Only apply if Status is Planned ('P')
                            if (loadedStatus == "P")
                            {
                                // IF DECREASING QUANTITY
                                if (newPlannedQty < currentMatrixRows)
                                {
                                    // Check only the rows that are about to be removed
                                    // Example: Current=5, New=3. We check rows 5 and 4.
                                    for (int i = currentMatrixRows; i > newPlannedQty; i--)
                                    {
                                        string chassisInRow = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(i).Specific).Value.Trim();

                                        // If ANY of the excess rows has data, BLOCK the action
                                        if (!string.IsNullOrEmpty(chassisInRow))
                                        {
                                            Utilities.Application.SBO_Application.SetStatusBarMessage(
                                                $"Error: Cannot decrease quantity to {newPlannedQty}. Row {i} contains a Chassis No. Please manually delete the row first.",
                                                BoMessageTime.bmt_Short, true);

                                            BubbleEvent = false;
                                            return;
                                        }
                                    }
                                }
                            }
                            // Status 'R' logic (existing)
                            else if (loadedStatus == "R" && newPlannedQty != currentMatrixRows)
                            {
                                Utilities.Application.SBO_Application.SetStatusBarMessage(
                                    "Error: Cannot change Planned Qty while Production Order is Released.",
                                    BoMessageTime.bmt_Short, true);
                                BubbleEvent = false;
                                return;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utilities.Application.SBO_Application.SetStatusBarMessage("Validation Error: " + ex.Message);
                    }
                }

                #region checkbox in chasis (open form and then select)
                //if (pVal.EventType == BoEventTypes.et_KEY_DOWN && pVal.ItemUID == "mtxEngChas" && pVal.ColUID == "colCha" && pVal.CharPressed == 9)
                //{
                //    string itemCode = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("ItemCode", 0).Trim();
                //    string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                //    double.TryParse(plannedQtyStr, out double plannedQty);

                //    Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                //    oMatrix.FlushToDataSource(); // Ensure we have latest data

                //    int filledRows = 0;
                //    for (int i = 1; i <= oMatrix.RowCount; i++)
                //    {
                //        string engineVal = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                //        if (!string.IsNullOrEmpty(engineVal))
                //        {
                //            filledRows++;
                //        }
                //    }

                //    double remainingQty = plannedQty - filledRows;

                //    if (remainingQty <= 0)
                //    {
                //        Utilities.Application.SBO_Application.SetStatusBarMessage("Planned Quantity limit reached. Cannot select more chassis.", BoMessageTime.bmt_Short, true);
                //        BubbleEvent = false;
                //        return;
                //    }

                //    // 4. Open Child Form
                //    this._Object = new ClsChasisSelection();
                //    this.OpenChildForm(Constants.Forms.ChasisSelection, false);

                //    // 5. Pass Parameters
                //    if (this._Object is ClsChasisSelection childForm)
                //    {
                //        childForm.ParentItemCode = itemCode;
                //        childForm.ParentFormUID = this.Form.UniqueID;
                //        childForm.ParentRowIndex = pVal.Row;
                //        childForm.RemainingQty = remainingQty; // Pass the limit

                //        childForm.FormDefault();
                //    }

                //    //this._Object = new ClsChasisSelection();
                //    //this.OpenChildForm(Constants.Forms.ChasisSelection, false);
                //    ////Utilities.LoadForm(ref this._Object, Constants.Forms.BatchSelection);
                //    ////((ClsChasisSelection)_Object).FormDefault();
                //    //if (this._Object is ClsChasisSelection childForm)
                //    //{
                //    //    childForm.ParentItemCode = itemCode; // Pass the value
                //    //    childForm.FormDefault(); // Now call the method that uses the value
                //    //}

                //    BubbleEvent = false;
                //}
                #endregion
            }
            #endregion

            #region After Action
            // After Action
            else
            {
               
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST)
                {
                    // First, handle the CFL event for your matrix
                    if (pVal.ItemUID == "mtxEngChas")
                    {
                        HandleCFLSelection(pVal);
                    }
                    //else if (pVal.ItemUID == "6" && this.Form.Mode == BoFormMode.fm_ADD_MODE)
                    //{
                    //    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    //    DataTable oDataTable = oCFLEvento.SelectedObjects;

                    //    // Always disable the tab first. It will be enabled only if all conditions are met.
                    //    this.Form.Items.Item("fldEngChas").Enabled = false;

                    //    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    //    {
                    //        try
                    //        {
                    //            // 1. Get ItemCode and Status
                    //            itemCode = oDataTable.GetValue(0, 0).ToString().Trim();
                    //            status = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                    //            // 2. Perform the validation checks
                    //            if (status == "P" && IsItemSerialManaged(itemCode))
                    //            {
                    //                // 3. If all checks pass, enable the tab
                    //                this.Form.Items.Item("fldEngChas").Enabled = true;
                    //                Utilities.Application.SBO_Application.SetStatusBarMessage("Serial item selected. Engine/Chassis selection is now available.", BoMessageTime.bmt_Short, false);
                    //            }
                    //            else if (status != "P")
                    //            {
                    //                Utilities.Application.SBO_Application.SetStatusBarMessage("Tab cannot be enabled because Production Order status is not 'Planned'.", BoMessageTime.bmt_Short, true);
                    //            }
                    //            else
                    //            {
                    //                Utilities.Application.SBO_Application.SetStatusBarMessage("Item is not managed by serial numbers. Engine/Chassis selection is not available.", BoMessageTime.bmt_Short, true);
                    //            }
                    //        }
                    //        catch (Exception ex)
                    //        {
                    //            Utilities.Application.SBO_Application.MessageBox("Error after Item CFL selection: " + ex.Message);
                    //        }
                    //    }
                    //}
                }

                //Open Engine/Chasis Selection Tab only when Item is serial managed 
                if (pVal.EventType == BoEventTypes.et_VALIDATE && pVal.ItemUID == "6" && pVal.ItemChanged)
                {
                    string itemCode = ((SAPbouiCOM.EditText)this.Form.Items.Item("6").Specific).Value;
                    EnableEngineChassisTab(itemCode);

                }



                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && pVal.ItemUID == "6" && this.Form.Mode == BoFormMode.fm_ADD_MODE)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    DataTable oDataTable = oCFLEvento.SelectedObjects;

                    string itemCode;

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        try
                        {
                            itemCode = oDataTable.GetValue(0, 0).ToString().Trim();
                            System.Threading.Tasks.Task.Run(() =>
                            {

                                System.Threading.Thread.Sleep(200); // Wait for CFL to close

                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)this.Form.Items.Item("mtxEngChas").Specific;

                                // 1. CLEAR VISUAL ROWS
                                oMatrix.Clear();

                                // 2. CLEAR THE DATA SOURCES (The "Memory")
                                // If you don't do this, AddRow() will just bring back the old values!
                                string[] udsToClear = { "colSrNo", "colCha", "colEng", "colSet", "colTrans", "colLotNo", "colEgChNo", "colPinCod", "colStatus", "colPInvt", "colRProd", "colTTime", "colBatch", "colProdS", "colDel" };

                                foreach (string udsName in udsToClear)
                                {
                                    try
                                    {
                                        this.Form.DataSources.UserDataSources.Item(udsName).ValueEx = "";
                                    }
                                    catch { /* Handle if UDS doesn't exist */ }
                                }

                                // 3. NOW ADD THE NEW BLANK ROW
                                oMatrix.AddRow();

                                // Set the Serial Number for the new row
                                this.Form.DataSources.UserDataSources.Item("colSrNo").ValueEx = "1";


                                EnableEngineChassisTab(itemCode);

                                Recordset rsOCN = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                string queryOCN = $"SELECT  T1.\"Name\"  FROM \"OITM\" T0 INNER JOIN \"@OCNMASTER\" T1 ON T0.\"U_OCN\" = T1.\"Code\" WHERE T0.\"ItemCode\" = '{itemCode}'";
                                rsOCN.DoQuery(queryOCN);

                                if (rsOCN.RecordCount > 0)
                                {
                                    string ocnValue = rsOCN.Fields.Item("Name").Value.ToString();
                                    SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.Item(FormUID);

                                    //this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_OCNCode", 0, ocnValue);
                                    try
                                    {
                                        //oForm.Items.Item("txtOCN").Enabled = true;
                                        //((SAPbouiCOM.EditText)oForm.Items.Item("txtOCN").Specific).Value = ocnValue;
                                        //oForm.Items.Item("6").Click();
                                        //oForm.Items.Item("txtOCN").Enabled = false;
                                        this.Form.DataSources.UserDataSources.Item("udsOCN").ValueEx = ocnValue;
                                    }
                                    catch (Exception ex)
                                    {
                                        // Fallback: If UI update fails, refreshing the form usually forces 
                                        // the bound EditText to show the DataSource value.
                                    }
                                }
                            });

                        }

                        catch (Exception ex)
                        {
                            Utilities.Application.SBO_Application.SetStatusBarMessage("Error on Item Change: " + ex.Message);
                        }
                    }

                  

                }


                if (pVal.EventType==BoEventTypes.et_KEY_DOWN && pVal.ItemUID=="12" && !pVal.BeforeAction && pVal.CharPressed==9)
                {
                    plannedQty(pVal);
                }

                #region Button Pressed (Inventory Transfer Request)
                if (pVal.ItemUID == "btnCustom" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.ProgressBar pgBar = null;
                    try
                    {
                        pgBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Inventory Transfer Request... Please wait.", 1, false);
                        this.Form.Freeze(true);
                        // Utilities.Application.SBO_Application.SetStatusBarMessage("Creating Inventory Transfer Request... Please wait.", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                        string prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                        // Call your logic to create the transfer
                        TransferReqDocEntry = Create_Inventory_Transfer_Request(prodDocEntry);

                        if (TransferReqDocEntry > 0)
                        {
                            // InvtTransferNo
                            SAPbobsCOM.ProductionOrders oProdOrder = null;
                            try
                            {
                                oProdOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                                if (oProdOrder.GetByKey(Convert.ToInt32(prodDocEntry)))
                                {
                                    // Update the UDF
                                    oProdOrder.UserFields.Fields.Item("U_InvtTransferReq").Value = TransferReqDocEntry;

                                    // Commit changes
                                    int ret = oProdOrder.Update();

                                    if (ret != 0)
                                    {
                                        Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                                        Utilities.Application.SBO_Application.SetStatusBarMessage("Error linking Transfer No to Production Order: " + errMsg, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                    }
                                    else
                                    {
                                        // Update the UI immediately without needing to refresh the whole form
                                        try
                                        {
                                            this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_InvtTransferReq", 0, TransferReqDocEntry.ToString());
                                        }
                                        catch { /* Field might not be visible/bound, ignore UI update error */ }

                                        Utilities.Application.SBO_Application.SetStatusBarMessage($"Inventory Transfer Request {TransferReqDocEntry} created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Utilities.Application.SBO_Application.SetStatusBarMessage("Exception updating Header: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            finally
                            {
                                if (oProdOrder != null)
                                {
                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdOrder);
                                    oProdOrder = null;
                                }
                                //this.Form.Freeze(false);

                                //if (pgBar != null)
                                //{
                                //    pgBar.Stop();
                                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(pgBar);
                                //}
                                //this.Form.Items.Item("btnCustom").Enabled = false;

                                //try
                                //{
                                //    this.Form.Select();
                                //    if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                                //    {
                                //        Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                                //    }
                                //}
                                //catch (Exception ex)
                                //{
                                //}


                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        if (pgBar != null) { pgBar.Stop(); System.Runtime.InteropServices.Marshal.ReleaseComObject(pgBar); pgBar = null; }

                        Utilities.Application.SBO_Application.MessageBox("Error creating Request: " + ex.Message);
                    }
                    finally
                    {
                        if (pgBar != null)
                        {
                            try
                            {
                                pgBar.Stop();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pgBar);
                            }
                            catch { /* Ignore errors during cleanup */ }
                            pgBar = null; 
                        }
                        try
                        {
                            this.Form.Freeze(false);
                            this.Form.Items.Item("btnCustom").Enabled = false;

                            this.Form.Select();
                            if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                            {
                                Utilities.Application.SBO_Application.ActivateMenuItem("1304"); // Refresh
                            }
                        }
                        catch { /* If the form closed or refreshed, ignore UI errors */ }
                    }

                }
                #endregion

                //#region Button Pressed (Inventory Tranfer)
                //if (pVal.ItemUID == "btnInvtT" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                //{
                //        SAPbouiCOM.ProgressBar pgBar1 = null;

                //        pgBar1 = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Inventory Transfer and Work Order Details... Please wait.", 1, false);
                //        this.Form.Freeze(true);

                //        string prodDocEntryStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                //        string transferRequest = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_InvtTransferReq", 0).Trim();
                //        int prodDocEntry = Convert.ToInt32(prodDocEntryStr);
                //        int workOrderEntry;

                //    //SAPbobsCOM.ProductionOrders oProdOrder1 = null;
                //    Task.Run(() => {
                //        try
                //        {
                //            //STEP 1: Create Inventory Transfer
                //            //pgBar1.Text = "Creating Inventory Transfer...";
                //            int transferDocEntry = Create_Transfer_From_Request(Convert.ToInt32(transferRequest));
                //            //int transferDocEntry = 493;

                //            if (transferDocEntry <= 0)
                //            {
                //                //Utilities.Application.SBO_Application.StatusBar.SetText("Failed to create Inventory Transfer. Process halted.");
                //                return; // Exit early
                //            }
                //            else
                //            {
                //                // Utilities.Application.SBO_Application.StatusBar.SetText($"Step 1 Success: Inventory Transfer {transferDocEntry} created.");

                //                // STEP 2: Create Work Order Details (UDO)
                //                //pgBar1.Text = "Generating Work Order Details...";
                //                string diError = "";
                //                workOrderEntry = AddWorkOrderDetailsDI(prodDocEntry, out diError, transferDocEntry);

                //                if (workOrderEntry <= 0)
                //                {
                //                    //Utilities.Application.SBO_Application.SetStatusBarMessage($"Inventory Transfer {transferDocEntry} was created, but Work Order Details failed: {diError}", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                //                    // Note: Transfer is already created in DB, we should still try to link it if possible, 
                //                    // but per your requirement, we show a separate error here.
                //                    return;
                //                }
                //                // Utilities.Application.SBO_Application.StatusBar.SetText($"Step 2 Success: Work Order Details {workOrderEntry} generated.");
                //            }

                //        // STEP 3: Update Production Order with both values
                //        //pgBar.Text = "Updating Production Order...";
                //        SAPbobsCOM.ProductionOrders oProdOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                //            if (oProdOrder.GetByKey(prodDocEntry))
                //            {
                //                oProdOrder.UserFields.Fields.Item("U_InvtTransfer").Value = transferDocEntry;
                //                oProdOrder.UserFields.Fields.Item("U_WorkOrderEntry").Value = workOrderEntry;

                //                if (oProdOrder.Update() == 0)
                //                {
                //                // Update UI Fields
                //                try
                //                {
                //                    this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_InvtTransfer", 0, transferDocEntry.ToString());
                //                    this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_WorkOrderEntry", 0, workOrderEntry.ToString());
                //                }
                //                catch { }

                //                //Utilities.Application.SBO_Application.StatusBar.SetText("All steps completed successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                //                this.Form.Items.Item("btnInvtT").Enabled = false;

                //                //
                //                if (this.Form.Mode == BoFormMode.fm_UPDATE_MODE) this.Form.Mode = BoFormMode.fm_OK_MODE;

                //                Utilities.Application.SBO_Application.SetStatusBarMessage($"Inventory Transfer and Work Order Details created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, false);

                //                }
                //                else
                //                {
                //                    string sboErr = Utilities.Application.Company.GetLastErrorDescription();
                //                    Utilities.Application.SBO_Application.SetStatusBarMessage($"Transfer & Work Order created, but failed to link to Production Order: {sboErr}");
                //                }
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //            Utilities.Application.SBO_Application.SetStatusBarMessage("General Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                //        }
                //        finally
                //        {
                //            //if (oProdOrder1 != null)
                //            //{
                //            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdOrder1);
                //            //    oProdOrder1 = null;
                //            //}
                //        this.Form.Freeze(false);

                //        if (pgBar1 != null)
                //        {
                //            pgBar1.Stop();
                //            System.Runtime.InteropServices.Marshal.ReleaseComObject(pgBar1);
                //        }
                //        try
                //        {
                //            this.Form.Select();
                //            if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                //            {
                //                Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                //            }
                //        }
                //        catch (Exception ex)
                //        {
                //        }
                //    }
                //    });

                //}
                //#endregion

                #region Button Pressed (Inventory Transfer)
                if (pVal.ItemUID == "btnInvtT" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.ProgressBar pgBar1 = null;

                    // Capture UI values BEFORE starting the background task
                    string prodDocEntryStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                    string transferRequest = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_InvtTransferReq", 0).Trim();
                    int prodDocEntry = Convert.ToInt32(prodDocEntryStr);

                    pgBar1 = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Inventory Transfer and Work Order Details... Please wait.", 1, false);
                    this.Form.Freeze(true);

                    Task.Run(() => {
                        try
                        {
                            // STEP 1: Create Inventory Transfer
                            int transferDocEntry = Create_Transfer_From_Request(Convert.ToInt32(transferRequest));

                            if (transferDocEntry <= 0)
                            {
                                // This should already be handled by the exception inside Create_Transfer_From_Request
                                throw new Exception("Inventory Transfer creation returned an invalid ID.");
                            }

                            // STEP 2: Create Work Order Details (UDO)
                            string diError = "";
                            int workOrderEntry = AddWorkOrderDetailsDI(prodDocEntry, out diError, transferDocEntry);

                            if (workOrderEntry <= 0)
                            {
                                // Throw the specific error returned by the DI function
                                throw new Exception("Work Order Details Generation Failed: " + diError);
                            }

                            // STEP 3: Update Production Order with both values
                            SAPbobsCOM.ProductionOrders oProdOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);

                            if (oProdOrder.GetByKey(prodDocEntry))
                            {
                                oProdOrder.UserFields.Fields.Item("U_InvtTransfer").Value = transferDocEntry;
                                oProdOrder.UserFields.Fields.Item("U_WorkOrderEntry").Value = workOrderEntry;

                                if (oProdOrder.Update() == 0)
                                {
                                    // Update UI Fields (Safe because we aren't creating new COM objects, just setting values)
                                    try
                                    {
                                        this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_InvtTransfer", 0, transferDocEntry.ToString());
                                        this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_WorkOrderEntry", 0, workOrderEntry.ToString());
                                    }
                                    catch { }

                                    this.Form.Items.Item("btnInvtT").Enabled = false;
                                    if (this.Form.Mode == BoFormMode.fm_UPDATE_MODE) this.Form.Mode = BoFormMode.fm_OK_MODE;

                                    Utilities.Application.SBO_Application.SetStatusBarMessage($"Inventory Transfer {transferDocEntry} and Work Order {workOrderEntry} created successfully.", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                }
                                else
                                {
                                    // Capture DI API Error from the Update attempt
                                    throw new Exception("Failed to link to Production Order: " + Utilities.Application.Company.GetLastErrorDescription());
                                }
                            }
                            else
                            {
                                throw new Exception("Could not find Production Order " + prodDocEntry + " to update.");
                            }
                        }
                        catch (Exception ex)
                        {
                            // STOP the progress bar immediately on error
                            if (pgBar1 != null) { pgBar1.Stop(); }

                            // Show the exact error in RED in the status bar
                            //Utilities.Application.SBO_Application.SetStatusBarMessage("ERROR: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, true);

                            Utilities.Application.SBO_Application.MessageBox("Process Failed!\n\n" + ex.Message);
                            Utilities.Application.SBO_Application.SetStatusBarMessage("Error: " + ex.Message, BoMessageTime.bmt_Long, true);
                        }
                        finally
                        {
                            this.Form.Freeze(false);
                            if (pgBar1 != null)
                            {
                                pgBar1.Stop();
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(pgBar1);
                                pgBar1 = null;
                            }

                            // UI Cleanup/Refresh
                            try
                            {
                                this.Form.Select();
                                if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                                {
                                    Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                                }
                            }
                            catch { }
                        }
                    });
                }
                #endregion

                #region Push to Inventory text clicked (Creation of Receipt from Production)
                if (pVal.EventType == BoEventTypes.et_CLICK && pVal.ItemUID == "mtxEngChas" && pVal.ColUID == "colPInvt" && pVal.Row > 0 && !pVal.BeforeAction)
                {
                    try
                    {
                        SAPbobsCOM.Recordset rs = null;
                        Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                        string cellValue = ((EditText)oMatrix.Columns.Item("colPInvt").Cells.Item(pVal.Row).Specific).Value.Trim();

                        // Only proceed if the link text is "Push to Inventory"
                        if (cellValue == "Push to Inventory")
                        {
                            // Gather Data from Matrix Row
                            string prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                            string chassisNo = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(pVal.Row).Specific).Value.Trim();
                            string engineNo = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(pVal.Row).Specific).Value.Trim();
                            string keySet = ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(pVal.Row).Specific).Value.Trim();
                            string transNo = ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(pVal.Row).Specific).Value.Trim();
                            string lotNo = ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(pVal.Row).Specific).Value.Trim();

                            // Confirmation
                            int confirm = Utilities.Application.SBO_Application.MessageBox($"Create Receipt from Production for Chassis {chassisNo}?", 1, "Yes", "No");

                            if (confirm == 1)
                            {
                                SAPbouiCOM.ProgressBar progBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Creating Receipt from Production........", 1, false);
                                // Create SAP Receipt from Production Document
                                int receiptDocEntry = Create_ReceiptFrom_Production(Convert.ToInt32(prodDocEntry), chassisNo, engineNo, keySet, transNo);

                                if (receiptDocEntry > 0)
                                {
                                    // Update Database (Status and Link to Receipt)
                                    UpdateRowAfterReceipt(prodDocEntry, chassisNo, receiptDocEntry);

                                    UpdateProductionQuantities(prodDocEntry);

                                    // Update UI Visuals immediately
                                    ((EditText)oMatrix.Columns.Item("colStatus").Cells.Item(pVal.Row).Specific).Value = "Manufactured";
                                    ((EditText)oMatrix.Columns.Item("colPInvt").Cells.Item(pVal.Row).Specific).Value = "";
                                    ((EditText)oMatrix.Columns.Item("colRProd").Cells.Item(pVal.Row).Specific).Value = receiptDocEntry.ToString();

                                    Utilities.Application.SBO_Application.SetStatusBarMessage("Receipt from Production created successfully.", BoMessageTime.bmt_Short, false);

                                    rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    int finalBatchToSave = 0;
                                    Recordset rsBatchLookup = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    rsBatchLookup.DoQuery($@"SELECT ""U_BatchNo"" FROM ""@LOTNOMASTER"" WHERE ""Code"" = '{lotNo}'");

                                    //if (rsBatchLookup.RecordCount > 0)
                                    //{
                                    //    int masterCurrentBatch = Convert.ToInt32(rsBatchLookup.Fields.Item("U_BatchNo").Value);
                                    //    finalBatchToSave = masterCurrentBatch + 1;
                                    //    if (!string.IsNullOrEmpty(lotNo))
                                    //    {
                                    //        rs.DoQuery($@"UPDATE ""@LOTNOMASTER"" SET ""U_BatchNo"" = {finalBatchToSave} WHERE ""Code"" = '{lotNo}'");
                                    //    }

                                    //}
                                    //Utilities.Application.SBO_Application.SetStatusBarMessage("Batch number updated successfully in Lot Number Master.", BoMessageTime.bmt_Short, false);

                                }
                                progBar.Stop();
                                try
                                {
                                    this.Form.Select();
                                    if (Utilities.Application.SBO_Application.Menus.Item("1304").Enabled)
                                    {
                                        Utilities.Application.SBO_Application.ActivateMenuItem("1304");
                                    }
                                }
                                catch (Exception ex)
                                {
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Utilities.Application.SBO_Application.SetStatusBarMessage("Error on Push to Inventory: " + ex.Message);
                    }
                }
                #endregion

                // Inside After Action (!pVal.BeforeAction)
                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED)
                {
                    // Check if the user clicked ANY folder (Tab)
                    // Common SAP Folder UIDs: 106 (Components), 107 (General), 108 (Summary)
                    // Or check by ItemType
                    SAPbouiCOM.Item oItem = this.Form.Items.Item(pVal.ItemUID);

                    if (oItem.Type == BoFormItemTypes.it_FOLDER)
                    {
                        // SAP just enabled everything because you switched tabs.
                        // Now, force it back to disabled.
                        this.Form.Items.Item("txtOCN").Enabled = false;
                    }
                }

                if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "fldEngChas" && !pVal.BeforeAction)
                {
                    SetButtonStates();
                }

                // inventory transfer and request button ebable/disable after status changed
                if (pVal.EventType == BoEventTypes.et_COMBO_SELECT && pVal.ItemUID == "10")
                {
                    SetButtonStates();
                }

            }
            #endregion
        }
        #endregion

        #region HandleCFLSelection
        private void HandleCFLSelection(ItemEvent pVal)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
            DataTable oDataTable = oCFLEvento.SelectedObjects;
            // User cancelled the CFL
            if (oDataTable == null)
            {
                return;
            }

            try
            {
                this.Form.Freeze(true);
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                if (pVal.ColUID == "colCha")
                {
                    string selectedEngineNo = oDataTable.GetValue("DistNumber", 0).ToString();

                    // Set the selected engine number in the matrix
                    ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(pVal.Row).Specific).Value = selectedEngineNo;

                    if (!string.IsNullOrEmpty(selectedEngineNo))
                    {
                        Recordset rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        string query = $"SELECT T0.\"U_EngineNo\", T0.\"U_SetKey\", T0.\"U_TransNo\", T1.\"U_LotNo\",T1.\"DocEntry\", T0.\"U_PinCode\", T1.\"U_BatchNo\"  FROM \"@ENGCHASISMMC\" T0 INNER JOIN \"@ENGCHASISMMH\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE \"U_ChasisNo\" = '{selectedEngineNo.Replace("'", "''")}'";
                        rs.DoQuery(query);

                        if (rs.RecordCount > 0)
                        {
                            // Get the values from the recordset
                            string engineNo = rs.Fields.Item("U_EngineNo").Value.ToString();
                            string keySet = rs.Fields.Item("U_SetKey").Value.ToString();
                            string transNo = rs.Fields.Item("U_TransNo").Value.ToString();
                            string LotNo = rs.Fields.Item("U_LotNo").Value.ToString();
                            string docEntry = rs.Fields.Item("DocEntry").Value.ToString();
                            string pinCode = rs.Fields.Item("U_PinCode").Value.ToString();
                            string batchNo = rs.Fields.Item("U_BatchNo").Value.ToString();

                            int currentBatch = 0;
                            int.TryParse(rs.Fields.Item("U_BatchNo").Value.ToString(), out currentBatch);

                            // Set the retrieved values in the corresponding columns of the matrix
                            //((EditText)oMatrix.Columns.Item("colCha").Cells.Item(pVal.Row).Specific).Value = engineNo;
                            ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(pVal.Row).Specific).Value = engineNo;
                            ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(pVal.Row).Specific).Value = keySet;
                            ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(pVal.Row).Specific).Value = transNo;
                            ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(pVal.Row).Specific).Value = LotNo;
                            ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(pVal.Row).Specific).Value = docEntry;
                            ((EditText)oMatrix.Columns.Item("colPinCod").Cells.Item(pVal.Row).Specific).Value = pinCode;
                            ((EditText)oMatrix.Columns.Item("colBatch").Cells.Item(pVal.Row).Specific).Value = (currentBatch + 1).ToString();
                            ((EditText)oMatrix.Columns.Item("colStatus").Cells.Item(pVal.Row).Specific).Value = "Pending";
                            ((EditText)oMatrix.Columns.Item("colDel").Cells.Item(pVal.Row).Specific).Value = "Delete";

                            //for Production Series No
                            //Model
                            string itemCode = ((SAPbouiCOM.EditText)this.Form.Items.Item("6").Specific).Value;
                            string modelPart = "";
                            int hyphenIndex = itemCode.IndexOf('-');
                            if (hyphenIndex > 0)
                                modelPart = itemCode.Substring(0, hyphenIndex);
                            else
                                modelPart = itemCode.Length >= 4 ? itemCode.Substring(0, 4) : itemCode;

                            //LotNo
                            string lotPart = "";
                            if (!string.IsNullOrEmpty(LotNo) && LotNo.Length >= 2)
                                lotPart = "00" + LotNo.Substring(LotNo.Length - 2);
                            else
                                lotPart = "00" + LotNo;

                            // OCN Code
                            string ocnCode = ((SAPbouiCOM.EditText)this.Form.Items.Item("txtOCN").Specific).Value;

                            // StartDate (Month/Year)
                            string startDateStr = ((SAPbouiCOM.EditText)this.Form.Items.Item("234000007").Specific).Value;
                            string monthPart = "";
                            string yearPart = "";

                            
                            if (!string.IsNullOrEmpty(startDateStr) && startDateStr.Length == 8)
                            {
                                DateTime dt = DateTime.ParseExact(startDateStr, "yyyyMMdd", null);
                                monthPart = dt.ToString("MMM").ToUpper(); 
                                yearPart = dt.ToString("yy");            
                            }

                            // BatchNo
                            string batchPart = "";
                            if (string.IsNullOrEmpty((currentBatch + 1).ToString()) || (currentBatch + 1).ToString().Equals("0"))
                                batchPart = "01";
                            else
                                batchPart = (currentBatch + 1).ToString().Trim().PadLeft(2, '0');

                            //CONCATENATE ALL
                            string prodSeries = $"{modelPart}{lotPart}{ocnCode}{monthPart}{yearPart}{batchPart}";
                            ((EditText)oMatrix.Columns.Item("colProdS").Cells.Item(pVal.Row).Specific).Value = prodSeries;

                        }

                        // Clean up the recordset object
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                        rs = null;
                    }
                }
                else
                {
                    //cfl for others if needed (chasis, setkey, transmission column) 
                    string selectedValue = oDataTable.GetValue("DistNumber", 0).ToString();
                    ((EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Value = selectedValue;
                }

                if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                {
                    this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                }

                // Add a new empty row if the current row is the last one.
                //if (pVal.Row == oMatrix.RowCount)
                //{
                //    // check if the last row has a value in the engine column before adding a new row
                //    string engineNoInLastRow = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(oMatrix.RowCount).Specific).Value;
                //    if (!string.IsNullOrEmpty(engineNoInLastRow))
                //    {
                //        oMatrix.AddRow();
                //    }
                //}

                //if (pVal.Row == oMatrix.RowCount)
                //{
                //    string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                //    double plannedQty = 0;
                //    double.TryParse(plannedQtyStr, out plannedQty); 

                //    int currentFilledRows = 0;
                //    for (int i = 1; i <= oMatrix.RowCount; i++)
                //    {
                //        string engineInRow = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                //        if (!string.IsNullOrEmpty(engineInRow))
                //        {
                //            currentFilledRows++;
                //        }
                //    }

                //    string engineNoInLastRow = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(oMatrix.RowCount).Specific).Value.Trim();
                //    if (!string.IsNullOrEmpty(engineNoInLastRow) && currentFilledRows < plannedQty)
                //    {
                //        oMatrix.AddRow(); // Only add a new row if we are below the limit.
                //    }
                //}
                //string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                //double plannedQty = 0;
                //double.TryParse(plannedQtyStr, out plannedQty); 

                //oMatrix.FlushToDataSource(); 
                //int currentFilledRows = 0;
                //for (int i = 1; i <= oMatrix.RowCount; i++)
                //{
                //    string engineInRow = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                //    if (!string.IsNullOrEmpty(engineInRow))
                //    {
                //        currentFilledRows++;
                //    }
                //}

                //// Only add a new row if the user just filled the LAST row AND the total number of filled rows
                //// is now LESS THAN the planned quantity.
                //if (pVal.Row == oMatrix.RowCount && currentFilledRows < plannedQty)
                //{
                //    oMatrix.AddRow(); // Only add a new row if we are still under the limit.

                //    int newRowIndex = oMatrix.RowCount;

                //    // Optional but good practice: set the new serial number.
                //    //((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(newRowIndex).Specific).Value = newRowIndex.ToString();

                //    // Clear all other data-entry columns in the new row
                //    ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //    ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //    ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //    ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //    ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //    ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(newRowIndex).Specific).Value = string.Empty;
                //}
                plannedQty(pVal);

            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Error after CFL selection: " + ex.Message);
            }
            finally
            {
                this.Form.Freeze(false);
            }
        }
        #endregion

        #region FormData_Event
        public override void FormData_Event(ref BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            base.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);

            if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD && BusinessObjectInfo.BeforeAction)
            {
                // Reset the form for the next "Add" operation by disabling the tab.
                this.Form.Items.Item("fldEngChas").Enabled = false;
            }

            // This event runs AFTER a form loads existing data.
            if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD && !BusinessObjectInfo.BeforeAction)
            {
                //this.Form.Freeze(true);
                //Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                //oMatrix.Clear();

                //string loadedDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();

                //if (!string.IsNullOrEmpty(loadedDocEntry))
                //{
                //    UpdateProductionQuantities(loadedDocEntry);
                //}

                this.Form.Items.Item("fldEngChas").Enabled = false;
                string loadedItemCode = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("ItemCode", 0).Trim();
                string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                // Enable the tab ONLY if the loaded item is valid for your functionality.
                if (IsItemSerialManaged(loadedItemCode))
                {
                    this.Form.Items.Item("fldEngChas").Enabled = true;
                }
                else
                {
                    this.Form.Items.Item("fldEngChas").Enabled = false;
                }
                if (loadedStatus != "C")
                {
                    //this.Form.Items.Item("btnCustom").Visible = false;
                    LoadMatrixData();
                }
                SetButtonStates();
                if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                {
                    this.Form.Items.Item("U_InvtTransfer").Enabled = false;
                    this.Form.Items.Item("U_InvtTransferReq").Enabled = false;
                    this.Form.Items.Item("U_WorkOrderEntry").Enabled = false;
                    this.Form.Items.Item("txtOCN").Enabled = false;
                }
            }

            if ((BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE) && BusinessObjectInfo.BeforeAction)
            {
                string validationMessage = "";
                string Status = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                string prodDoc = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                if (Status == "P")
                {
                    if (!ValidateEngineCount(out validationMessage))
                    {
                        //Utilities.Application.SBO_Application.MessageBox(validationMessage);
                        Utilities.Application.SBO_Application.MessageBox(validationMessage);
                        BubbleEvent = false;
                        return;
                    }
                    if (!ValidateCanSave(out validationMessage))
                    {
                        Utilities.Application.SBO_Application.MessageBox("Validation Failed: " + validationMessage);
                        BubbleEvent = false;
                        return;
                    }
                }
            }

            // after clicking add or update button in sap's standard form
            if ((BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD || BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_UPDATE))
            {
                if (BusinessObjectInfo.ActionSuccess && !BusinessObjectInfo.BeforeAction)
                {
                    //SAPbouiCOM.ProgressBar pb = null;
                    SAPbouiCOM.Form oForm = Utilities.Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                    try
                    {
                        this.Form.Freeze(true);
                        string prodDocEntry = "";

                        if (BusinessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_ADD)
                        {
                            prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                        }
                        else // For an UPDATE event
                        {
                            // For an EXISTING record, the DocEntry is on the form.
                            prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                        }
                        //prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();

                        System.Threading.Tasks.Task.Run(() =>
                        {
                            try
                            {
                                // Give SAP 500ms to finish its internal save and UI state change
                                System.Threading.Thread.Sleep(500);

                                // Re-calculate and update quantities using the SQL Method 
                                if (!string.IsNullOrEmpty(prodDocEntry))
                                {
                                    // Run your database updates
                                    SaveMatrixData(prodDocEntry);
                                    UpdateProductionQuantities(prodDocEntry);
                                }                                
                                this.Form.Close();
                            }
                            catch (Exception ex)
                            {
                                // Silently catch errors in background thread to prevent crash popups
                            }
                        });
                    }
                    catch (Exception ex)
                    {
                        Utilities.Application.SBO_Application.MessageBox("Error in FormDataEvent: " + ex.Message);
                    }
                    //finally
                    //{
                    //    if (pb != null)
                    //    {
                    //        pb.Stop();
                    //        System.Runtime.InteropServices.Marshal.ReleaseComObject(pb);
                    //        pb = null;
                    //    }

                    //    // Safety check: if the form failed to close, unfreeze it
                    //    try { this.Form.Freeze(false); } catch { }
                    //}
                }
            }
        }
        #endregion

        #region Validation Logic - before adding Production Order
        private bool ValidateCanSave(out string errorMessage)
        {
            errorMessage = "";

            //string currentLotNo = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_LotNo", 0).Trim();
            //if (string.IsNullOrEmpty(currentLotNo))
            //{
            //    errorMessage = "Lot Number is missing in the Production Order Header.";
            //    return false;
            //}
            Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
            oMatrix.FlushToDataSource(); 

            int rowCount = oMatrix.VisualRowCount;
            if (rowCount == 0) return true;

            for (int i = 1; i <= rowCount; i++)
            {
                string engineNo = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                string chasisNo = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(i).Specific).Value.Trim();
                string masterDocEntry = ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(i).Specific).Value.Trim();

                if (!string.IsNullOrEmpty(chasisNo))
                {
                    for (int j = i + 1; j <= rowCount; j++)
                    {
                        string nextEngine = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(j).Specific).Value.Trim();
                        if (chasisNo == nextEngine)
                        {
                            errorMessage = $"Duplicate Chasis No '{chasisNo}' found in row {i} and {j}.";
                            return false;
                        }
                    }

                    //string currentProdDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                    //string excludeClause = "";
                    //if (!string.IsNullOrEmpty(currentProdDocEntry) && currentProdDocEntry != "0")
                    //{
                    //    excludeClause = $" AND \"U_ProdOrdNo\" <> '{currentProdDocEntry}'";
                    //}

                    //string query = $"SELECT \"DocEntry\", \"U_ProdOrdNo\" FROM \"@ENGCHASISMMC\" WHERE \"U_ChasisNo\" = '{engineNo.Replace("'", "''")}' AND \"U_Status\" = 'WIP' {excludeClause}";

                    //Recordset rsCheck = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                    //rsCheck.DoQuery(query);
                    //if (rsCheck.RecordCount > 0)
                    //{
                    //    string otherOrder = rsCheck.Fields.Item("U_ProdOrdNo").Value.ToString();
                    //    errorMessage = $"Engine '{engineNo}' is already assigned to Production Order #{otherOrder}. Please select a different engine.";
                    //    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheck);
                    //    return false;
                    //}
                    //System.Runtime.InteropServices.Marshal.ReleaseComObject(rsCheck);
                }
            }

            return true;
        }
        #endregion

        #region Data Loading
        private void LoadMatrixData()
        {
            if (this.Form.Mode == BoFormMode.fm_ADD_MODE)
            {
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                oMatrix.Clear();
                oMatrix.AddRow();
                //empty rows should not have delete 
                ((EditText)oMatrix.Columns.Item("colDel").Cells.Item(1).Specific).Value = "";
                return;
            }

            bool wasInOkMode = (this.Form.Mode == BoFormMode.fm_OK_MODE);

            try
            {
                this.Form.Freeze(true);
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                oMatrix.Clear();

                string prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();

                if (string.IsNullOrEmpty(prodDocEntry)) return;

                //OCN COde
                SAPbobsCOM.Recordset rsOCNLoad = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                try
                {
                    // Fetch the UDF value directly from OWOR
                    string qry = $"SELECT \"U_OCNCode\" FROM OWOR WHERE \"DocEntry\" = '{prodDocEntry}'";
                    rsOCNLoad.DoQuery(qry);
                    if (rsOCNLoad.RecordCount > 0)
                    {
                        string val = rsOCNLoad.Fields.Item("U_OCNCode").Value.ToString();
                        // Set the UserDataSource - This updates the screen instantly
                        this.Form.DataSources.UserDataSources.Item("udsOCN").ValueEx = val;
                    }
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsOCNLoad);
                }

                Recordset rs1 = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query1 = $"SELECT  T1.\"U_BatchNo\"  FROM \"@ENGCHASISMMC\" T0 INNER JOIN \"@ENGCHASISMMH\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"U_ProdOrdNo\" = '{prodDocEntry}'";
                rs1.DoQuery(query1);

                Recordset rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                string query = $"SELECT * FROM \"@ENGCHASPO\" WHERE \"U_ProdDocEntry\" = {prodDocEntry} ORDER BY \"U_SrNo\" ";
                rs.DoQuery(query);

                string batchNo = "";
                if (rs1.RecordCount > 0)
                {
                    batchNo = rs1.Fields.Item("U_BatchNo").Value.ToString();
                }

                if (rs.RecordCount > 0)
                {
                    //oMatrix.AddRow(rs.RecordCount);

                    for (int i = 0; i < rs.RecordCount; i++)
                    {
                        //while (!rs.EoF) { 
                        oMatrix.AddRow();
                        //int matrixRow = oMatrix.VisualRowCount;
                        int matrixRow = i + 1;
                        string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                        //string prodS = rs.Fields.Item("U_ProdSeriesNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_SrNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_EngineNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_ChasisNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_KeySet").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_TransNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_LotNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_EnChNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colPinCod").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_PinCode").Value.ToString();
                        //((EditText)oMatrix.Columns.Item("colStatus").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_Status").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colPInvt").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_PushToInvt").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colRProd").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_ReceiptFrmProd").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colTTime").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_TotalTime").Value.ToString();
                        //((EditText)oMatrix.Columns.Item("colBatch").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_BatchNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colBatch").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_BatchNo").Value.ToString();
                        ((EditText)oMatrix.Columns.Item("colProdS").Cells.Item(matrixRow).Specific).Value = rs.Fields.Item("U_ProdSeriesNo").Value.ToString();
                        //((EditText)oMatrix.Columns.Item("colDel").Cells.Item(matrixRow).Specific).Value = "Delete";

                        string lineStatus = "";
                        try { lineStatus = rs.Fields.Item("U_Status").Value.ToString().Trim(); } catch { }


                       ((EditText)oMatrix.Columns.Item("colStatus").Cells.Item(matrixRow).Specific).Value = lineStatus;
                        //string chassisNo = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(matrixRow).Specific).Value.Trim();

                        if (lineStatus == "Completed")
                        {
                            // Case: Completed
                            // Show "Push to Inventory"
                            SAPbouiCOM.EditText oLinkText = (EditText)oMatrix.Columns.Item("colPInvt").Cells.Item(matrixRow).Specific;
                            oLinkText.Value = "Push to Inventory";

                            // Style: Blue and Underlined
                            oLinkText.ForeColor = 16711680; // Standard SAP Link Blue
                            oLinkText.TextStyle = (int)BoTextStyle.ts_UNDERLINE;
                        }

                        SAPbouiCOM.EditText oEditDel = (EditText)oMatrix.Columns.Item("colDel").Cells.Item(matrixRow).Specific;

                        if (loadedStatus == "R")
                        {
                            oEditDel.Value = "Delete";
                            oEditDel.ForeColor = 8421504;
                        }
                        else
                        {
                            oEditDel.Value = "Delete";
                            // Keep color Red (255)
                            oEditDel.ForeColor = 255;
                        }



                        rs.MoveNext();
                    }
                }

                string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                double plannedQty = 0;
                double.TryParse(plannedQtyStr, out plannedQty);
                string docStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                if (docStatus == "P" && oMatrix.VisualRowCount < plannedQty)
                {
                    oMatrix.AddRow();
                }
                FreezeMatrixRows();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Error loading matrix data: " + ex.Message);
            }
            finally
            {
                if (wasInOkMode) this.Form.Mode = BoFormMode.fm_OK_MODE;
                SetButtonStates();
                this.Form.Freeze(false);
            }
        }

        #endregion

        #region Save Matrix to database table
        private void SaveMatrixData(string prodDocEntry)
        {
            SAPbobsCOM.UserTable oUserTable = null;
            SAPbobsCOM.Recordset rs = null;
            bool transactionStarted = false;

            try
            {
                string currentStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                // START TRANSACTION
                if (!Utilities.Application.Company.InTransaction)
                {
                    Utilities.Application.Company.StartTransaction();
                    transactionStarted = true;
                }

                rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                oUserTable = Utilities.Application.Company.UserTables.Item("ENGCHASPO");



                // Release existing items in Master Data (Set to Available)
                // We do this to clear the state before re-applying the current matrix state
                //string releaseQuery = "UPDATE \"@ENGCHASISMMC\" " +
                //                      "SET \"U_ProdOrdNo\" = NULL, \"U_Status\" = 'Available' " +
                //                      $"WHERE \"U_ProdOrdNo\" = '{prodDocEntry}'";
                //rs.DoQuery(releaseQuery);

                /////////////// batch no 

                //Recordset rsExist = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                //string existingBatchQuery = $@"SELECT TOP 1 T0.""U_BatchNo"" FROM ""@ENGCHASPO"" T0 INNER JOIN OWOR T1 ON T0.""U_ProdDocEntry"" = T1.""DocEntry""
                //               WHERE T0.""U_ProdDocEntry"" = '{prodDocEntry}' 
                //               AND T0.""U_BatchNo"" IS NOT NULL AND T0.""U_BatchNo"" <> '0' AND T0.""U_BatchNo"" <> '' and T1.""Status"" = 'R'";
                //rsExist.DoQuery(existingBatchQuery);

                //string batchCheck = rsExist.Fields.Item("U_BatchNo").Value.ToString();
                //bool isBatchAlreadyAssigned = batchCheck != "";
                ////bool isBatchAlreadyAssigned = rsExist.RecordCount > 0;
                //string globalBatchNo = isBatchAlreadyAssigned ? rsExist.Fields.Item("U_BatchNo").Value.ToString() : "";
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(rsExist);
                /////////////////////

                //  Delete existing Child Table rows
                string deleteQuery = $"DELETE FROM \"@ENGCHASPO\" WHERE \"U_ProdDocEntry\" = '{prodDocEntry}'";
                rs.DoQuery(deleteQuery);

                string ocnCode = this.Form.DataSources.UserDataSources.Item("udsOCN").ValueEx;
                UpdateUdfViaDI(Convert.ToInt32(prodDocEntry), ocnCode);


                // Iterate Matrix and Insert
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string engineNo = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                    string chasisNo = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(i).Specific).Value.Trim();
                    string set = ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(i).Specific).Value.Trim();
                    string transNo = ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(i).Specific).Value.Trim();
                    string masterDocEntry = ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(i).Specific).Value.Trim();
                    string lotNo = ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(i).Specific).Value.Trim();
                    string srNoStr = ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(i).Specific).Value;
                    string pinCode = ((EditText)oMatrix.Columns.Item("colPinCod").Cells.Item(i).Specific).Value;
                    string status = ((EditText)oMatrix.Columns.Item("colStatus").Cells.Item(i).Specific).Value;
                    string pushToInvt = ((EditText)oMatrix.Columns.Item("colPInvt").Cells.Item(i).Specific).Value;
                    string receiptFrmProd = ((EditText)oMatrix.Columns.Item("colRProd").Cells.Item(i).Specific).Value;
                    string totalTime = ((EditText)oMatrix.Columns.Item("colTTime").Cells.Item(i).Specific).Value;
                    string batchNo = ((EditText)oMatrix.Columns.Item("colBatch").Cells.Item(i).Specific).Value;
                    string prodSeriesNo = ((EditText)oMatrix.Columns.Item("colProdS").Cells.Item(i).Specific).Value;

                    if (!string.IsNullOrEmpty(engineNo))
                    {
                        //update batch no
                        int finalBatchToSave = 0;
                        int.TryParse(batchNo, out finalBatchToSave);

                            //if (currentStatus == "R" && !string.IsNullOrEmpty(masterDocEntry))
                            //{
                            //    Recordset rsBatchLookup = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                            //    rsBatchLookup.DoQuery($@"SELECT ""U_BatchNo"" FROM ""@ENGCHASISMMH"" WHERE ""DocEntry"" = '{masterDocEntry}'");

                            //    if (rsBatchLookup.RecordCount > 0)
                            //    {
                            //        int masterCurrentBatch = Convert.ToInt32(rsBatchLookup.Fields.Item("U_BatchNo").Value);
                            //        int prodOrderBatch = finalBatchToSave;

                            //        if (masterCurrentBatch == prodOrderBatch)
                            //        {
                            //            finalBatchToSave = masterCurrentBatch + 1;
                            //        }
                            //        else
                            //        {
                            //            finalBatchToSave = prodOrderBatch;
                            //        }

                            //        // 1. Update Master Tables
                            //        rs.DoQuery($@"UPDATE ""@ENGCHASISMMH"" SET ""U_BatchNo"" = {finalBatchToSave} WHERE ""DocEntry"" = '{masterDocEntry}'");
                            //        if (!string.IsNullOrEmpty(lotNo))
                            //        {
                            //            rs.DoQuery($@"UPDATE ""@LOTNOMASTER"" SET ""U_BatchNo"" = {finalBatchToSave} WHERE ""Code"" = '{lotNo}'");
                            //        }

                            //        //production Order Series
                            //        string itemCodeHeader = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("ItemCode", 0).Trim();
                            //        string modelPart = itemCodeHeader.Contains("-") ? itemCodeHeader.Split('-')[0] : (itemCodeHeader.Length >= 4 ? itemCodeHeader.Substring(0, 4) : itemCodeHeader);
                            //        string lotPart = (!string.IsNullOrEmpty(lotNo) && lotNo.Length >= 2) ? "00" + lotNo.Substring(lotNo.Length - 2) : "00" + lotNo;

                            //        string startDateStr = ((SAPbouiCOM.EditText)this.Form.Items.Item("234000007").Specific).Value;
                            //        string monthPart = "", yearPart = "";
                            //        if (!string.IsNullOrEmpty(startDateStr) && startDateStr.Length == 8)
                            //        {
                            //            DateTime dt = DateTime.ParseExact(startDateStr, "yyyyMMdd", null);
                            //            monthPart = dt.ToString("MMM").ToUpper();
                            //            yearPart = dt.ToString("yy");
                            //        }

                            //        string batchPart = finalBatchToSave.ToString().PadLeft(2, '0');
                            //        prodSeriesNo = $"{modelPart}{lotPart}{ocnCode}{monthPart}{yearPart}{batchPart}";
                            //    }
                            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(rsBatchLookup);
                            //}




                            string uniqueCode = $"{prodDocEntry}-{i}";

                        if (oUserTable.GetByKey(uniqueCode))
                        {
                            oUserTable.Remove();
                        }

                        oUserTable.Code = uniqueCode;
                        oUserTable.Name = uniqueCode;

                        int.TryParse(srNoStr, out int srNo);

                        oUserTable.UserFields.Fields.Item("U_ProdDocEntry").Value = prodDocEntry;
                        oUserTable.UserFields.Fields.Item("U_SrNo").Value = srNo;
                        oUserTable.UserFields.Fields.Item("U_EngineNo").Value = engineNo;
                        oUserTable.UserFields.Fields.Item("U_ChasisNo").Value = chasisNo;
                        oUserTable.UserFields.Fields.Item("U_KeySet").Value = set;
                        oUserTable.UserFields.Fields.Item("U_TransNo").Value = transNo;
                        oUserTable.UserFields.Fields.Item("U_LotNo").Value = lotNo;
                        oUserTable.UserFields.Fields.Item("U_EnChNo").Value = masterDocEntry;
                        oUserTable.UserFields.Fields.Item("U_PinCode").Value = pinCode;
                        oUserTable.UserFields.Fields.Item("U_Status").Value = status;
                        oUserTable.UserFields.Fields.Item("U_PushToInvt").Value = pushToInvt;
                        oUserTable.UserFields.Fields.Item("U_ReceiptFrmProd").Value = receiptFrmProd;
                        oUserTable.UserFields.Fields.Item("U_TotalTime").Value = totalTime;
                        oUserTable.UserFields.Fields.Item("U_BatchNo").Value = Convert.ToString(finalBatchToSave);
                        oUserTable.UserFields.Fields.Item("U_ProdSeriesNo").Value = prodSeriesNo;

                        if (oUserTable.Add() != 0)
                        {
                            throw new Exception($"Failed to save Row {i}: {Utilities.Application.Company.GetLastErrorDescription()}");
                        }

                        // 4. Update Master Data to Locked/WIP
                        if (!string.IsNullOrEmpty(masterDocEntry))
                        {
                            string lockQuery = "UPDATE \"@ENGCHASISMMC\" " +
                                               $"SET \"U_ProdOrdNo\" = '{prodDocEntry}', \"U_Status\" = 'WIP' " +
                                               $"WHERE \"DocEntry\" = '{masterDocEntry}'";
                            if (!string.IsNullOrEmpty(engineNo))
                                lockQuery += $" AND \"U_EngineNo\" = '{engineNo.Replace("'", "''")}'";
                            if (!string.IsNullOrEmpty(chasisNo))
                                lockQuery += $" AND \"U_ChasisNo\" = '{chasisNo.Replace("'", "''")}'";
                            if (!string.IsNullOrEmpty(set))
                                lockQuery += $" AND \"U_SetKey\" = '{set.Replace("'", "''")}'";
                            if (!string.IsNullOrEmpty(transNo))
                                lockQuery += $" AND \"U_TransNo\" = '{transNo.Replace("'", "''")}'";

                            rs.DoQuery(lockQuery);
                            //string updateQuery = $@"
                            //    UPDATE ""@ENGCHASISMMH""
                            //    SET ""U_BatchNo"" = (
                            //        SELECT COUNT(DISTINCT ""U_ProdOrdNo"")
                            //        FROM ""@ENGCHASISMMC""
                            //        WHERE ""DocEntry"" = '{masterDocEntry}'
                            //    )
                            //    WHERE ""DocEntry"" = '{masterDocEntry}' ";

                            //rs.DoQuery(updateQuery);

                            //string updateBatchQuery = $@"
                            //    UPDATE ""@LOTNOMASTER""
                            //    SET ""U_BatchNo"" = (
                            //        SELECT COUNT(DISTINCT ""U_ProdOrdNo"")
                            //        FROM ""@ENGCHASISMMC""
                            //        WHERE ""DocEntry"" = '{masterDocEntry}'
                            //    )
                            //    WHERE ""Code"" = '{lotNo}' ";

                            //rs.DoQuery(updateBatchQuery);

                        }
                    }
                }
                // COMMIT TRANSACTION
                if (transactionStarted)
                {
                    Utilities.Application.Company.EndTransaction(BoWfTransOpt.wf_Commit);
                }

                this.Form.Freeze(true);
                try
                {
                    oMatrix.Clear();    
                    //oMatrix.AddRow(); 
                }
                finally
                {
                    this.Form.Freeze(false);
                }
            }
            catch (Exception ex)
            {
                // ROLLBACK TRANSACTION
                if (transactionStarted && Utilities.Application.Company.InTransaction)
                {
                    Utilities.Application.Company.EndTransaction(BoWfTransOpt.wf_RollBack);
                }

                Utilities.Application.SBO_Application.MessageBox("Error saving Engine Data. Changes rolled back. Details: " + ex.Message);
            }
            finally
            {
                if (oUserTable != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                if (rs != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
            }
        }
        #endregion

        #region Freeze Matrix Rows
        //private void FreezeMatrixRows()
        //{
        //    try
        //    {
        //        this.Form.Freeze(true);

        //        string docStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
        //        Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;

        //        bool isFrozen = (docStatus == "R" || docStatus == "L" || docStatus == "C");

        //        oMatrix.Columns.Item("colCha").Editable = !isFrozen;

        //        int color = isFrozen ? 15790320 : -1; // 15790320 is Light Gray, -1 is Default/White

        //        frozenRows.Clear();

        //        //for (int i = 1; i <= oMatrix.RowCount; i++)
        //        //{
        //        //    if (isFrozen)
        //        //    {
        //        //        frozenRows.Add(i);
        //        //    }

        //        //    for (int c = 1; c < oMatrix.Columns.Count; c++)
        //        //    {
        //        //        oMatrix.CommonSetting.SetCellBackColor(i, c, color);
        //        //    }
        //        //}
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.MessageBox("Error in FreezeMatrixRows: " + ex.Message);
        //    }
        //    finally
        //    {
        //        this.Form.Freeze(false);
        //    }
        //}

        private void FreezeMatrixRows()
        {
            try
            {
                this.Form.Freeze(true);

                string docStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;

                bool isFrozen = (docStatus == "R" || docStatus == "L" || docStatus == "C");

                for (int i = 0; i < oMatrix.Columns.Count; i++)
                {
                    SAPbouiCOM.Column oCol = oMatrix.Columns.Item(i);

                    if (isFrozen)
                    {
                        oCol.Editable = false;
                    }
                    else
                    {
                        if (oCol.UniqueID == "colCha")
                        {
                            oCol.Editable = true;
                        }
                        else
                        {
                            oCol.Editable = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Utilities.Application.SBO_Application.MessageBox("Error in FreezeMatrixRows: " + ex.Message);
            }
            finally
            {
                this.Form.Freeze(false);
            }
        }
        #endregion

        #region Item Serial Managed
        private bool IsItemSerialManaged(string selectedItemCode)
        {
            // Return false immediately if the item code is empty
            if (string.IsNullOrEmpty(selectedItemCode))
            {
                return false;
            }

            Recordset rs = null;
            try
            {
                rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                // Prepare the query to prevent SQL injection
                string query = $"SELECT \"ManSerNum\" FROM OITM WHERE \"ItemCode\" = '{selectedItemCode.Replace("'", "''")}'";
                rs.DoQuery(query);

                // Check if the query returned a result and if the value is 'Y'
                if (rs.RecordCount > 0 && rs.Fields.Item("ManSerNum").Value.ToString() == "Y")
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Error checking item serial management: " + ex.Message);
                return false; // Assume false on error
            }
            finally
            {
                // Always release COM objects
                if (rs != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);
                    rs = null;
                }
            }

            // the item was not found or is not serial managed
            return false;
        }
        #endregion

        #region Validation for Planned Qty and No of engine Chasis Selected
        private bool ValidateEngineCount(out string errorMessage)
        {
            errorMessage = "";
            try
            {
                string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                double plannedQty = 0;
                double.TryParse(plannedQtyStr, out plannedQty);

                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;
                int currentFilledRows = 0;

                oMatrix.FlushToDataSource();

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    string engineInRow = ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(i).Specific).Value.Trim();
                    if (!string.IsNullOrEmpty(engineInRow))
                    {
                        currentFilledRows++;
                    }
                }
                if (currentFilledRows != plannedQty)
                {
                    errorMessage = $"The number of selected Engine/Chassis numbers ({currentFilledRows}) must exactly match the Planned Quantity ({plannedQty}).";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = "An unexpected error occurred during engine count validation: " + ex.Message;
                return false;
            }
        }
        #endregion

        #region Planned Qty
        public void plannedQty(ItemEvent pVal)
        {
            try
            {
                this.Form.Freeze(true);

                // Get Planned Qty
                string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                double.TryParse(plannedQtyStr, out double plannedQty);

                string loadedStatus = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;

                //in header qty changed
                if (loadedStatus == "P")
                {
                    if (pVal.ItemUID == "12")
                    {

                        if (oMatrix.VisualRowCount < plannedQty)
                        {
                            int rowsMissing = (int)plannedQty - oMatrix.VisualRowCount;

                            //add all the remaining rows
                            for (int i = 0; i < rowsMissing; i++)
                            {
                                oMatrix.AddRow();
                                int newRowIndex = oMatrix.VisualRowCount;

                                ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(newRowIndex).Specific).Value = newRowIndex.ToString();

                                // Clear Data
                                ClearRowData(oMatrix, newRowIndex);
                            }
                        }
                        //else if (oMatrix.VisualRowCount > plannedQty)
                        //{
                        //    //for (int i = oMatrix.VisualRowCount; i > plannedQty; i--)
                        //    //{
                        //    //    oMatrix.DeleteRow(i);
                        //    //}
                        //    //Utilities.Application.SBO_Application.MessageBox(
                        //    //  $"Error: Planned Quantity ({plannedQty}) cannot be less than the current number of rows ({oMatrix.VisualRowCount}).\n\nPlease delete the excess rows in the matrix manually first, then update the Planned Quantity.");
                        //}
                        else if (oMatrix.VisualRowCount > plannedQty)
                        {
                            for (int i = oMatrix.VisualRowCount; i > plannedQty; i--)
                            {
                                oMatrix.DeleteRow(i);
                            }
                        }

                    }

                    //in matrix
                    else
                    {
                        oMatrix.FlushToDataSource();

                        int currentFilledRows = 0;
                        for (int i = 1; i <= oMatrix.RowCount; i++)
                        {
                            // Check if engine/chassis is filled
                            string engineInRow = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(i).Specific).Value.Trim();
                            if (!string.IsNullOrEmpty(engineInRow))
                            {
                                currentFilledRows++;
                            }
                        }

                        // Only add if user is on the LAST row AND we are still under the limit
                        if (pVal.Row == oMatrix.RowCount && currentFilledRows < plannedQty)
                        {
                            oMatrix.AddRow();
                            int newRowIndex = oMatrix.RowCount;

                            ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(newRowIndex).Specific).Value = newRowIndex.ToString();
                            ClearRowData(oMatrix, newRowIndex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error in plannedQty: " + ex.Message);
            }
            finally
            {
                this.Form.Freeze(false);
            }
        }
        #endregion

        #region Clear Row Data
        private void ClearRowData(Matrix oMatrix, int rowIndex)
        {
            ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colEng").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colSet").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colTrans").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colLotNo").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colEgChNo").Cells.Item(rowIndex).Specific).Value = string.Empty;
            ((EditText)oMatrix.Columns.Item("colDel").Cells.Item(rowIndex).Specific).Value = string.Empty;
        }
        #endregion

        #region Delete Specific Row and ReSequence serial no 
        private void DeleteSpecificRow(int rowToDelete)
        {
            this.Form.Freeze(true);
            try
            {
                string prodDocEntry = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("DocEntry", 0).Trim();
                Matrix oMatrix = (Matrix)this.Form.Items.Item("mtxEngChas").Specific;

                if (rowToDelete > 0 && rowToDelete <= oMatrix.RowCount)
                {
                    oMatrix.DeleteRow(rowToDelete);
                    if (!string.IsNullOrEmpty(prodDocEntry))
                    {
                        rs1 = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
                        //update the Engine/Chasis to Available from WIP if deleted the row
                        string releaseQuery = "UPDATE \"@ENGCHASISMMC\" " +
                                              "SET \"U_ProdOrdNo\" = NULL, \"U_Status\" = 'Available' " +
                                              $"WHERE \"U_ProdOrdNo\" = '{prodDocEntry}'";
                        rs1.DoQuery(releaseQuery);


                        //  Delete existing Child Table rows
                        string deleteQuery = $"DELETE FROM \"@ENGCHASPO\" WHERE \"U_ProdDocEntry\" = '{prodDocEntry}'";
                        rs1.DoQuery(deleteQuery);

                        if (this.Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            this.Form.Mode = BoFormMode.fm_UPDATE_MODE;
                        }
                    }



                    // Resequence
                    for (int i = 1; i <= oMatrix.RowCount; i++)
                    {
                        ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(i).Specific).Value = i.ToString();
                    }

                    // Add blank row if needed
                    string plannedQtyStr = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                    double plannedQty = 0;
                    double.TryParse(plannedQtyStr, out plannedQty);

                    if (oMatrix.VisualRowCount < plannedQty)
                    {
                        string lastRowEngine = "";
                        if (oMatrix.RowCount > 0)
                        {
                            lastRowEngine = ((EditText)oMatrix.Columns.Item("colCha").Cells.Item(oMatrix.RowCount).Specific).Value;
                        }

                        if (oMatrix.RowCount == 0 || !string.IsNullOrEmpty(lastRowEngine))
                        {
                            oMatrix.AddRow();
                            int newRowIndex = oMatrix.RowCount;
                            ClearRowData(oMatrix, newRowIndex);
                            ((EditText)oMatrix.Columns.Item("colSrNo").Cells.Item(newRowIndex).Specific).Value = newRowIndex.ToString();
                            // Ensure new row has no Delete text
                            ((EditText)oMatrix.Columns.Item("colDel").Cells.Item(newRowIndex).Specific).Value = "";
                        }
                    }

                    oMatrix.FlushToDataSource();
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Error deleting row: " + ex.Message);
            }
            finally
            {
                this.Form.Freeze(false);
            }
        }
        #endregion

        #region Auto Inventory Transfer Request Before Issue
        private int Create_Inventory_Transfer_Request(string ProdDocEntry)
        {
            int newTransferEntry = 0;

            try
            {
                int prodEntry = Convert.ToInt32(ProdDocEntry);

                // Fetch BOM Items
                SAPbobsCOM.Recordset rsBOM = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string bomQuery = $@"
            SELECT  
                T0.""ItemCode"",
                T0.""PlannedQty"",
                T0.""IssueType"",
                T0.""wareHouse"",
                T0.""BaseQty"" AS ""UnitReq"",
                T0.""LineNum"",
                T2.""ManSerNum"",
                T2.""ManBtchNum"",
                T1.""PlannedQty"" AS ""HeaderQty"", T2.""DfltWH"" , T1.""ItemCode"" AS ""parentItemCode""
            FROM WOR1 T0
            INNER JOIN OWOR T1 ON T0.""DocEntry"" = T1.""DocEntry""
            INNER JOIN OITM T2 ON T0.""ItemCode"" = T2.""ItemCode""
            WHERE T0.""DocEntry"" = '{prodEntry}' and T2.""InvntItem"" = 'Y' and  T2.""PrchseItem"" = 'Y' ";

                rsBOM.DoQuery(bomQuery);

                // Fetch chassis mapping
                SAPbobsCOM.Recordset rsMap = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rsMap.DoQuery($@"SELECT * FROM ""@ENGCHASPO"" WHERE ""U_ProdDocEntry""='{prodEntry}'");

                if (rsMap.EoF)
                    throw new Exception("Mapping not found for chassis: " + prodEntry);

                // Create Inventory Transfer
                SAPbobsCOM.StockTransfer oTransfer = (SAPbobsCOM.StockTransfer)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);
                

                oTransfer.DocDate = DateTime.Now;
                oTransfer.UserFields.Fields.Item("U_ProdEntry").Value = prodEntry;
                oTransfer.Comments = $"For the product: '{rsBOM.Fields.Item("parentItemCode").Value}' with planned Quantity: '{rsBOM.Fields.Item("HeaderQty").Value}'.";
                //oTransfer.FromWarehouse =  // MAIN SOURCE

                while (!rsBOM.EoF)
                {
                    string itemCode = rsBOM.Fields.Item("ItemCode").Value.ToString();
                    double plannedQty = Convert.ToDouble(rsBOM.Fields.Item("PlannedQty").Value);
                    double headerQty = Convert.ToDouble(rsBOM.Fields.Item("HeaderQty").Value);
                    string whTo = rsBOM.Fields.Item("wareHouse").Value.ToString();
                    string isSerial = rsBOM.Fields.Item("ManSerNum").Value.ToString();
                    string frmWhs = rsBOM.Fields.Item("DfltWH").Value.ToString();
                    //string isBatch = rsBOM.Fields.Item("ManBtchNum").Value.ToString();

                    double transferQty = plannedQty;

                    oTransfer.Lines.ItemCode = itemCode;
                    oTransfer.Lines.FromWarehouseCode = frmWhs;
                    oTransfer.Lines.WarehouseCode = whTo;
                    oTransfer.Lines.Quantity = transferQty;

                    // ================================
                    // SERIAL-MANAGED ITEMS
                    // ================================
                    //if (isSerial == "Y")
                    //{
                    //    // Serial mapping based on item category
                    //    SAPbobsCOM.Recordset rsSerType = (SAPbobsCOM.Recordset)
                    //        Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    //    rsSerType.DoQuery($@"SELECT ""U_SerialType"" FROM OITM WHERE ""ItemCode""='{itemCode}'");

                    //    string serialType = rsSerType.Fields.Item("U_SerialType").Value.ToString();
                    //    string serValue = "";

                    //    switch (serialType)
                    //    {
                    //        case "ENGINE": serValue = rsMap.Fields.Item("U_EngineNo").Value.ToString(); break;
                    //        case "CHASSIS": serValue = rsMap.Fields.Item("U_ChasisNo").Value.ToString(); break;
                    //        case "TRANSMISSION": serValue = rsMap.Fields.Item("U_TransNo").Value.ToString(); break;
                    //        case "KEY": serValue = rsMap.Fields.Item("U_KeySet").Value.ToString(); break;
                    //        default:
                    //            throw new Exception("Serial item type not mapped: " + itemCode);
                    //    }

                    //    oTransfer.Lines.SerialNumbers.InternalSerialNumber = serValue;
                    //    oTransfer.Lines.SerialNumbers.Quantity = 1;
                    //    oTransfer.Lines.SerialNumbers.Add();
                    //}

                    if (isSerial == "Y")
                    {
                        // 1. Determine the Serial Type for this BOM Item (Engine, Chassis, etc.)
                        SAPbobsCOM.Recordset rsSerType = (SAPbobsCOM.Recordset)
                            Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                        rsSerType.DoQuery($@"SELECT ""U_SerialType"" FROM OITM WHERE ""ItemCode""='{itemCode}'");

                        if (rsSerType.RecordCount > 0)
                        {
                            string serialType = rsSerType.Fields.Item("U_SerialType").Value.ToString().Trim();

                            // 2. IMPORTANT: Move to the first record of the mapping table before looping
                            if (rsMap.RecordCount > 0)
                            {
                                rsMap.MoveFirst();

                                // 3. Loop through all rows in the Mapping Recordset
                                while (!rsMap.EoF)
                                {
                                    string serValue = "";

                                    // Get the specific value based on the Item Type
                                    switch (serialType)
                                    {
                                        case "ENGINE":
                                            serValue = rsMap.Fields.Item("U_EngineNo").Value.ToString();
                                            break;
                                        case "CHASSIS":
                                            serValue = rsMap.Fields.Item("U_ChasisNo").Value.ToString();
                                            break;
                                        case "TRANSMISSION":
                                            serValue = rsMap.Fields.Item("U_TransNo").Value.ToString();
                                            break;
                                        case "KEY":
                                            serValue = rsMap.Fields.Item("U_KeySet").Value.ToString();
                                            break;
                                        // Add default or case to skip if type is not found
                                        default:
                                            serValue = "";
                                            break;
                                    }

                                    // 4. Add the Serial Number to the Transfer Line
                                    if (!string.IsNullOrEmpty(serValue))
                                    {
                                        oTransfer.Lines.SerialNumbers.InternalSerialNumber = serValue;
                                        // oTransfer.Lines.SerialNumbers.Quantity = 1; // Default is 1

                                        // Add the serial line
                                        oTransfer.Lines.SerialNumbers.Add();
                                    }

                                    // 5. Move to the next row in the map
                                    rsMap.MoveNext();
                                }
                            }
                        }

                        // Clean up
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSerType);
                    }

                    // ================================
                    // BATCH-MANAGED ITEMS
                    // ================================

                    #region Batch
                    //    if (isBatch == "Y")
                    //    {
                    //        SAPbobsCOM.Recordset rsBatch = (SAPbobsCOM.Recordset)
                    //            Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

                    //        string batchQuery = $@"
                    //    SELECT ""BatchNum"", ""Quantity"" 
                    //    FROM OIBT 
                    //    WHERE ""ItemCode"" = '{itemCode}'
                    //      AND ""WhsCode"" = 'WHS2'
                    //      AND ""Quantity"" > 0
                    //    ORDER BY ""InDate""
                    //";

                    //        rsBatch.DoQuery(batchQuery);

                    //        double qtyNeeded = transferQty;

                    //        while (!rsBatch.EoF && qtyNeeded > 0)
                    //        {
                    //            string batchNum = rsBatch.Fields.Item("BatchNum").Value.ToString();
                    //            double batchQty = Convert.ToDouble(rsBatch.Fields.Item("Quantity").Value);

                    //            double consumeQty = Math.Min(batchQty, qtyNeeded);

                    //            oTransfer.Lines.BatchNumbers.BatchNumber = batchNum;
                    //            oTransfer.Lines.BatchNumbers.Quantity = consumeQty;
                    //            oTransfer.Lines.BatchNumbers.Add();

                    //            qtyNeeded -= consumeQty;
                    //            rsBatch.MoveNext();
                    //        }
                    //    }
                    #endregion

                    oTransfer.Lines.Add();
                    rsBOM.MoveNext();
                }

                int ret = oTransfer.Add();
                if (ret != 0)
                {
                    Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                    throw new Exception($"Inventory Transfer DI Error ({errCode}): {errMsg}");
                }

                string key = Utilities.Application.Company.GetNewObjectKey();
                int.TryParse(key, out newTransferEntry);

                Utilities.Application.SBO_Application.StatusBar.SetText(
                    $"Inventory Transfer Created: {newTransferEntry}",
                    BoMessageTime.bmt_Short,
                    BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                //Utilities.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                //throw new Exception("Inventory Transfer Reuquest Failed: " + ex.Message);
                throw;
                //return 0;
            }

            return newTransferEntry;
        }
        #endregion

        //#region Auto Inventory Transfer 
        //public int Create_Transfer_From_Request(int requestDocEntry)
        //{
        //    int transferEntry = 0;

        //    // 1. Get the Source Request Object
        //    SAPbobsCOM.StockTransfer oReq = (SAPbobsCOM.StockTransfer)
        //        Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryTransferRequest);

        //    if (!oReq.GetByKey(requestDocEntry))
        //    {
        //        throw new Exception($"Request ID {requestDocEntry} not found.");
        //    }

        //    // 2. Initialize the Target Transfer Object
        //    SAPbobsCOM.StockTransfer oTrans = (SAPbobsCOM.StockTransfer)
        //        Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

        //    try
        //    {
        //        // --- Copy Header Details ---
        //        oTrans.CardCode = oReq.CardCode; // If BP is associated
        //        oTrans.DocDate = DateTime.Now;
        //        oTrans.TaxDate = DateTime.Now;
        //        oTrans.Comments = "Copied from Request " + requestDocEntry + ". " + oReq.Comments;

        //        // Note: FromWarehouse is usually defined on the lines for Transfers, 
        //        // but if header level usage is required by your settings:
        //        // oTrans.FromWarehouse = oReq.FromWarehouse; 

        //        // --- Loop through Request Lines to build Transfer Lines ---
        //        for (int i = 0; i < oReq.Lines.Count; i++)
        //        {
        //            oReq.Lines.SetCurrentLine(i);

        //            // Add a new line to Transfer (except for the first default line)
        //            if (i > 0) oTrans.Lines.Add();

        //            // Link to Base Document ("Copy To" logic)
        //            oTrans.Lines.BaseType = SAPbobsCOM.InvBaseDocTypeEnum.InventoryTransferRequest; // 1250000001
        //            oTrans.Lines.BaseEntry = oReq.DocEntry;
        //            oTrans.Lines.BaseLine = oReq.Lines.LineNum;

        //            // Copy Line Details
        //            oTrans.Lines.ItemCode = oReq.Lines.ItemCode;
        //            oTrans.Lines.FromWarehouseCode = oReq.Lines.FromWarehouseCode;
        //            oTrans.Lines.WarehouseCode = oReq.Lines.WarehouseCode;
        //            oTrans.Lines.Quantity = oReq.Lines.Quantity;

        //            // --- Copy Serial Numbers (Crucial for Serial Items) ---
        //            // Even if linked via BaseEntry, you often must confirm which Serials are moving
        //            for (int s = 0; s < oReq.Lines.SerialNumbers.Count; s++)
        //            {
        //              c
        //              oReq.Lines.SerialNumbers.SetCurrentLine(s);

        //                // Add serial line to target
        //                if (s > 0) oTrans.Lines.SerialNumbers.Add();

        //                oTrans.Lines.SerialNumbers.InternalSerialNumber = oReq.Lines.SerialNumbers.InternalSerialNumber;
        //                // oTrans.Lines.SerialNumbers.Quantity = 1; // Usually implied
        //            }

        //            // --- Copy Batch Numbers (If applicable) ---
        //            for (int b = 0; b < oReq.Lines.BatchNumbers.Count; b++)
        //            {
        //                oReq.Lines.BatchNumbers.SetCurrentLine(b);

        //                if (b > 0) oTrans.Lines.BatchNumbers.Add();

        //                oTrans.Lines.BatchNumbers.BatchNumber = oReq.Lines.BatchNumbers.BatchNumber;
        //                oTrans.Lines.BatchNumbers.Quantity = oReq.Lines.BatchNumbers.Quantity;
        //            }
        //        }

        //        // 3. Add the Transfer Document
        //        int ret = oTrans.Add();

        //        if (ret != 0)
        //        {
        //            Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
        //            throw new Exception($"Failed to copy to Transfer: {errCode} - {errMsg}");
        //        }

        //        // Get the new DocEntry
        //        string key = Utilities.Application.Company.GetNewObjectKey();
        //        int.TryParse(key, out transferEntry);
        //    }
        //    catch (Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        // Cleanup COM objects
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oReq);
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrans);
        //    }

        //    return transferEntry;
        //}
        //#endregion


        //public int Create_Transfer_From_Request(int requestDocEntry)
        //{
        //    int transferEntry = 0;
        //    SAPbobsCOM.Recordset rsHeader = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbobsCOM.Recordset rsLines = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    SAPbobsCOM.StockTransfer oTrans = (SAPbobsCOM.StockTransfer)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

        //    try
        //    {
        //        // 1. Extract Header Data from OWTQ
        //        string sqlHeader = $@"SELECT ""CardCode"", ""Filler"", ""Comments"", ""DocDate"" FROM OWTQ WHERE ""DocEntry"" = {requestDocEntry}";
        //        rsHeader.DoQuery(sqlHeader);

        //        if (rsHeader.EoF)
        //        {
        //            throw new Exception($"Inventory Transfer Request {requestDocEntry} not found in database.");
        //        }

        //        // 2. Map Header to the new Transfer
        //        oTrans.CardCode = rsHeader.Fields.Item("CardCode").Value.ToString();
        //        oTrans.FromWarehouse = rsHeader.Fields.Item("Filler").Value.ToString();
        //        oTrans.Comments = $"Manual Transfer based on Request {requestDocEntry}. " + rsHeader.Fields.Item("Comments").Value.ToString();
        //        oTrans.DocDate = DateTime.Now;

        //        // 3. Extract Lines Data from WTQ1
        //        string sqlLines = $@"SELECT ""ItemCode"", ""Quantity"", ""FromWhsCod"", ""WhsCode"", ""LineNum"" FROM ""WTQ1"" WHERE ""DocEntry"" = {requestDocEntry}";
        //        rsLines.DoQuery(sqlLines);

        //        int lineCount = 0;
        //        while (!rsLines.EoF)
        //        {
        //            if (lineCount > 0) oTrans.Lines.Add();

        //            oTrans.Lines.ItemCode = rsLines.Fields.Item("ItemCode").Value.ToString();
        //            oTrans.Lines.Quantity = Convert.ToDouble(rsLines.Fields.Item("Quantity").Value);
        //            oTrans.Lines.FromWarehouseCode = rsLines.Fields.Item("FromWhsCod").Value.ToString();
        //            oTrans.Lines.WarehouseCode = rsLines.Fields.Item("WhsCode").Value.ToString();
        //            oTrans.Lines.BaseEntry = requestDocEntry;
        //            oTrans.Lines.BaseLine = Convert.ToInt32(rsLines.Fields.Item("LineNum").Value.ToString());
        //            oTrans.Lines.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;
        //            // IMPORTANT: Since we are NOT using "Copy To" logic (BaseEntry), 
        //            // the Request will remain OPEN unless you manually close it later.
        //            // No BaseEntry or BaseLine is set here.

        //            #region Optional: Manual Serial/Batch Extraction
        //            // If the items are Serial/Batch managed and you want to pull the 
        //            // SPECIFIC numbers selected in the Request:

                    
        //            SAPbobsCOM.Recordset rsSerials = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //            string sqlSerials = $@"SELECT T1.""DistNumber"" FROM ""SRI1"" T0 INNER JOIN ""OSRN"" T1 ON T0.""SysSerial"" = T1.""SysNumber""
        //                                  WHERE T0.""BaseEntry"" = {requestDocEntry} AND T0.""BaseType"" = 1250000001 
        //                                  AND T0.""LineNum"" = {rsLines.Fields.Item("LineNum").Value}";
        //            rsSerials.DoQuery(sqlSerials);

        //            int sCount = 0;
        //            while(!rsSerials.EoF)
        //            {
        //                if (sCount > 0) oTrans.Lines.SerialNumbers.Add();
        //                oTrans.Lines.SerialNumbers.InternalSerialNumber = rsSerials.Fields.Item("DistNumber").Value.ToString();
        //                sCount++;
        //                rsSerials.MoveNext();
        //            }
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSerials);
                    
        //            #endregion

        //            lineCount++;
        //            rsLines.MoveNext();
        //        }

        //        // 4. Add the Transfer
        //        int ret = oTrans.Add();

        //        if (ret != 0)
        //        {
        //            string errMsg = Utilities.Application.Company.GetLastErrorDescription();
        //            throw new Exception($"DI Error: {errMsg}");
        //        }

        //        // Success - Get new Key
        //        transferEntry = int.Parse(Utilities.Application.Company.GetNewObjectKey());
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($"Extraction Error: {ex.Message}");
        //    }
        //    finally
        //    {
        //        // Cleanup COM
        //        if (rsHeader != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsHeader);
        //        if (rsLines != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsLines);
        //        if (oTrans != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oTrans);
        //    }

        //    return transferEntry;
        //}

        #region Create Inventory Transfer from Inventory Transfer Request (With Serials)
        public int Create_Transfer_From_Request(int requestDocEntry)
        {
            int newTransferEntry = 0;

            SAPbobsCOM.Recordset rsLines = null;
            SAPbobsCOM.Recordset rsSerials = null;
            SAPbobsCOM.Recordset rsItem = null;
            SAPbobsCOM.StockTransfer oTransfer = null;
           
            try
            {
                // ---------------------------------
                // 1. GET REQUEST LINES
                // ---------------------------------
                rsLines = (SAPbobsCOM.Recordset)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rsLines.DoQuery($@"
            SELECT 
                T0.""ItemCode"",
                T0.""Quantity"",
                T0.""FromWhsCod"",
                T0.""WhsCode"",
                T0.""LineNum"",
                T1.""ManSerNum""
            FROM WTQ1 T0
            INNER JOIN OITM T1 ON T0.""ItemCode"" = T1.""ItemCode""
            WHERE T0.""DocEntry"" = {requestDocEntry}");

                if (rsLines.EoF)
                    throw new Exception("No lines found in Inventory Transfer Request.");

                // ---------------------------------
                // 2. CREATE INVENTORY TRANSFER
                // ---------------------------------
                oTransfer = (SAPbobsCOM.StockTransfer)
                    Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);

                oTransfer.DocDate = DateTime.Today;
                oTransfer.Comments = $"Auto Transfer from ITR {requestDocEntry}";

                int lineIndex = 0;

                // ---------------------------------
                // 3. LOOP REQUEST LINES
                // ---------------------------------
                while (!rsLines.EoF)
                {
                    if (lineIndex > 0)
                        oTransfer.Lines.Add();

                    string itemCode = rsLines.Fields.Item("ItemCode").Value.ToString();
                    double qty = Convert.ToDouble(rsLines.Fields.Item("Quantity").Value);
                    string fromWhs = rsLines.Fields.Item("FromWhsCod").Value.ToString();
                    string toWhs = rsLines.Fields.Item("WhsCode").Value.ToString();
                    int baseLine = Convert.ToInt32(rsLines.Fields.Item("LineNum").Value);
                    string isSerial = rsLines.Fields.Item("ManSerNum").Value.ToString();

                    oTransfer.Lines.ItemCode = itemCode;
                    oTransfer.Lines.Quantity = qty;
                    oTransfer.Lines.FromWarehouseCode = fromWhs;
                    oTransfer.Lines.WarehouseCode = toWhs;
                    oTransfer.Lines.BaseType = InvBaseDocTypeEnum.InventoryTransferRequest;
                    oTransfer.Lines.BaseEntry = requestDocEntry;
                    oTransfer.Lines.BaseLine = baseLine;

                    if (isSerial == "Y")
                    {
                        rsSerials = (SAPbobsCOM.Recordset)
                            Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        rsSerials.DoQuery($@"
                    SELECT distinct T1.""DistNumber"" FROM ""SRI1"" T0
                    INNER JOIN ""OSRN"" T1 ON T0.""SysSerial"" = T1.""SysNumber"" and T0.""ItemCode"" = T1.""ItemCode""
                    WHERE T0.""BaseEntry"" = {requestDocEntry} AND T0.""BaseType"" = '1250000001' AND T0.""BaseLinNum"" = {baseLine}");
                        int serialAdded = 0;

                        while (!rsSerials.EoF)
                        {
                            if (serialAdded > 0)
                                oTransfer.Lines.SerialNumbers.Add();

                            oTransfer.Lines.SerialNumbers.InternalSerialNumber =
                                rsSerials.Fields.Item("DistNumber").Value.ToString();

                           

                            serialAdded++;
                            rsSerials.MoveNext();
                        }

                        if (serialAdded != qty)
                            throw new Exception(
                                $"Serial count mismatch for item {itemCode}. Expected {qty}, Found {serialAdded}");
                    }

                    rsLines.MoveNext();
                    lineIndex++;
                }

                int ret = oTransfer.Add();
                if (ret != 0)
                {
                    Utilities.Application.Company.GetLastError(out int code, out string msg);
                    throw new Exception($"DI Error ({code}): {msg}");

                }

                newTransferEntry =
                    Convert.ToInt32(Utilities.Application.Company.GetNewObjectKey());
            }
            catch (Exception ex)
            {
                //throw new Exception("Inventory Transfer Failed: " + ex);
                throw;
            }
            finally
            {
                if (rsLines != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsLines);
                if (rsSerials != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsSerials);
                if (rsItem != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(rsItem);
                if (oTransfer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oTransfer);

                //if (hasError)
                //{
                //    Utilities.Application.SBO_Application.StatusBar.SetText(
                //        errorMsg,
                //        BoMessageTime.bmt_Long,
                //        BoStatusBarMessageType.smt_Error);
                //}
                //else
                //{
                //    Utilities.Application.SBO_Application.StatusBar.SetText(
                //        $"Inventory Transfer completed successfully.",
                //        BoMessageTime.bmt_Short,
                //        BoStatusBarMessageType.smt_Success);
                //}
            }

            return newTransferEntry;
        }
        #endregion

        #region Create Receipt From Production
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
        #endregion

        #region Update DB After Receipt
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

                // Update Master Table (@ENGCHASISMMC)
                // Set Status to Manufactured and save the Receipt DocEntry
                string queryMaster = $@"UPDATE ""@ENGCHASISMMC"" 
                                SET ""U_Status"" = 'Manufactured', 
                                    ""U_RepProd"" = '{receiptDocEntry}'
                                WHERE ""U_ChasisNo"" = '{chassisNo}' AND ""U_ProdOrdNo"" = '{prodDocEntry}'";
                rs.DoQuery(queryMaster);
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
        #endregion

        #region Inventory Transfer/Request (Button Enable/Disable)
        private void SetButtonStates()
        {
            try
            {
                if (this.Form == null) return;

                if (this.Form.Mode == BoFormMode.fm_ADD_MODE || this.Form.Mode == BoFormMode.fm_UPDATE_MODE)
                {
                    this.Form.Items.Item("btnCustom").Enabled = false;
                    this.Form.Items.Item("btnInvtT").Enabled = false;
                }

                string status = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();
                string reqDocNum = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_InvtTransferReq", 0).Trim();
                string transDocNum = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("U_InvtTransfer", 0).Trim();

                bool isReleased = (status == "R");
                bool isOkMode = (this.Form.Mode == BoFormMode.fm_OK_MODE);

                // Check if Request has a value (Not null/empty and not "0")
                bool hasRequest = !string.IsNullOrEmpty(reqDocNum) && reqDocNum != "0";

                // Check if Transfer has a value (Not null/empty and not "0")
                bool hasTransfer = !string.IsNullOrEmpty(transDocNum) && transDocNum != "0";

                // Logic for btnCustom (Request Button):
                // Enable ONLY if Released AND No Request yet AND No Transfer yet
                bool enableBtnCustom = isOkMode && isReleased && !hasRequest && !hasTransfer;

                // Logic for btnInvtT (Transfer Button):
                // Enable ONLY if Released AND Request exists AND No Transfer yet
                bool enableBtnInvtT = isOkMode && isReleased && hasRequest && !hasTransfer;

                    this.Form.Items.Item("btnCustom").Enabled = enableBtnCustom;
                    this.Form.Items.Item("btnInvtT").Enabled = enableBtnInvtT;

            }
            catch (Exception ex)
            {
                if (this.Form != null)
                {
                    this.Form.Items.Item("btnCustom").Enabled = false;
                    this.Form.Items.Item("btnInvtT").Enabled = false;
                }
                // Ideally log the error here
            }
        }
        #endregion

        #region Enable/Disable Engine/Chasis Tab 
        private void EnableEngineChassisTab(string itemCode)
        {
            try
            {
                this.Form.Freeze(true); 

                // Check Status 
                string status = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("Status", 0).Trim();

                // Logic Check
                bool shouldEnable = (status == "P" && IsItemSerialManaged(itemCode));

                // Enable engine/chasis Tab for selection
                this.Form.Items.Item("fldEngChas").Enabled = shouldEnable;

                if (shouldEnable)
                {
                    Utilities.Application.SBO_Application.SetStatusBarMessage("Serial item selected. Engine/Chassis selection is now available.", BoMessageTime.bmt_Short, false);
                }
                else
                {
                    // Utilities.Application.SBO_Application.SetStatusBarMessage("Engine/Chassis selection not available for this item.", BoMessageTime.bmt_Short, false);
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error enabling tab: " + ex.Message);
            }
            finally
            {
                this.Form.Freeze(false); 
            }
        }
        #endregion

        #region Update Production Quantities
        private void UpdateProductionQuantities(string docEntry)
        {
            SAPbobsCOM.Recordset oRec = null;
            SAPbobsCOM.ProductionOrders oProdOrder = null;
            SAPbobsCOM.Recordset rs = null;
            try
            {
                oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $@"
                    SELECT 
                        COUNT(T0.""Code"") AS ""CompQty"", 
                        ((T1.""PlannedQty"") - COUNT(T0.""Code"")) AS ""RemQty"" 
                    FROM ""@ENGCHASPO"" T0 
                    INNER JOIN OWOR T1 ON T0.""U_ProdDocEntry"" = T1.""DocEntry"" 
                    WHERE T0.""U_ProdDocEntry"" = '{docEntry}' AND T0.""U_Status"" = 'Manufactured' 
                    GROUP BY T1.""PlannedQty""";

                oRec.DoQuery(query);

                double completedQty = 0;
                double remainingQty = 0;

                if (oRec.RecordCount > 0)
                {
                    completedQty = Convert.ToDouble(oRec.Fields.Item("CompQty").Value);
                    remainingQty = Convert.ToDouble(oRec.Fields.Item("RemQty").Value);
                }
                else
                {
                    //if no Manufactured 
                    string sPlanned = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("PlannedQty", 0).Trim();
                    double.TryParse(sPlanned, out remainingQty);
                    completedQty = 0;
                }
                //this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_CompletedQty", 0, completedQty.ToString());
                //this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_RemainingQty", 0, remainingQty.ToString());

                //oProdOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                //if (oProdOrder.GetByKey(Convert.ToInt32(docEntry)))
                //{
                //    oProdOrder.UserFields.Fields.Item("U_CompletedQty").Value = completedQty.ToString();
                //    oProdOrder.UserFields.Fields.Item("U_RemainingQty").Value = remainingQty.ToString();

                //    int ret = oProdOrder.Update();
                //    if (ret != 0)
                //    {
                //        Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
                //        // Log error but don't stop the UI
                //    }
                //    else
                //    {
                //        //try
                //        //{
                //        //    this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_CompletedQty", 0, completedQty.ToString());
                //        //    this.Form.DataSources.DBDataSources.Item("OWOR").SetValue("U_RemainingQty", 0, remainingQty.ToString());

                //        //    Update the actual items on the form if they are not automatically binding
                //        //    if (this.Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                //        //        this.Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                //        //}
                //        //catch { /* Field might not be on the current form view */ }
                //    }
                //}
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdOrder);

                //Instead of using .update() use direct update query and then refreh the form

                rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string updateSql = $@"UPDATE OWOR SET ""U_CompletedQty"" = '{completedQty}', ""U_RemainingQty"" = '{remainingQty}' WHERE ""DocEntry"" = '{docEntry}'";
                rs.DoQuery(updateSql);

                // REFRESH THE UI
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.SetStatusBarMessage("Error calculating Qty: " + ex.Message);
            }
            finally
            {
                if (oRec != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
                    oRec = null;
                }
                if (oProdOrder != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdOrder);
            }
        }
        #endregion

        #region Create Work Order Details
        public int AddWorkOrderDetailsDI(int productionDocEntry, out string errorMessage, int transferNo)
        {
            errorMessage = "";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            // Suggestion: Declare these locally if they aren't class-level variables
            SAPbobsCOM.CompanyService oCompanyService;
            SAPbobsCOM.GeneralService oGeneralService;
            SAPbobsCOM.GeneralData oGeneralData;
            SAPbobsCOM.GeneralDataCollection oChildren;
            SAPbobsCOM.GeneralData oChild;

            try
            {
                oCompanyService = Utilities.Application.Company.GetCompanyService();
                oGeneralService = oCompanyService.GetGeneralService("WRKORDRDTLS");
                oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                SAPbobsCOM.GeneralDataParams oParams;
                oParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);

                // Header Query
                string queryHeader = $@"
                SELECT TOP 1 T0.""DocNum"", T0.""ItemCode"", T0.""ProdName"", T0.""PlannedQty"", 
                T1.""U_LotNo"", T1.""U_EnChNo"", T0.""U_InvtTransfer"", T2.""ItmsGrpNam""
                FROM OWOR T0  
                INNER JOIN ""@ENGCHASPO"" T1 ON T0.""DocEntry"" = T1.""U_ProdDocEntry"" 
                INNER JOIN OITM T3 ON T0.""ItemCode"" = T3.""ItemCode""
                INNER JOIN OITB T2 ON T3.""ItmsGrpCod"" = T2.""ItmsGrpCod""
                WHERE T0.""DocEntry"" = {productionDocEntry}"; 

                rs.DoQuery(queryHeader);
                if (rs.EoF) { errorMessage = "Production Order not found."; return -1; }

                // Set Header Fields
                oGeneralData.SetProperty("U_PrdOrdEnt", productionDocEntry.ToString());
                oGeneralData.SetProperty("U_ProdOrdNo", rs.Fields.Item("DocNum").Value.ToString());
                oGeneralData.SetProperty("U_ProductCode", rs.Fields.Item("ItemCode").Value.ToString());
                oGeneralData.SetProperty("U_ProductName", rs.Fields.Item("ProdName").Value.ToString());
                oGeneralData.SetProperty("U_PlannedQty", Convert.ToInt32(rs.Fields.Item("PlannedQty").Value));
                oGeneralData.SetProperty("U_InvtTransferNo", transferNo.ToString());
                oGeneralData.SetProperty("U_Model", rs.Fields.Item("ItmsGrpNam").Value.ToString());
                oGeneralData.SetProperty("U_Status", "Open");
                oGeneralData.SetProperty("U_PostDate", DateTime.Now);
                oGeneralData.SetProperty("U_UpdatedBy", Utilities.Application.Company.UserName);
                oGeneralData.SetProperty("U_DocNo", Utilities.getMaxColumnValueNum("@WRKORDRDTLSH", "U_DocNo"));

                oChildren = oGeneralData.Child("WRKORDRDTLSC");

                // Chassis Query
                string queryChassis = $@"SELECT ""U_EngineNo"", ""U_ChasisNo"", ""U_LotNo"", ""U_EnChNo"" 
                                 FROM ""@ENGCHASPO"" 
                                 WHERE ""U_ProdDocEntry"" = {productionDocEntry} AND ""U_EngineNo"" IS NOT NULL";

                SAPbobsCOM.Recordset rsChas = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsChas.DoQuery(queryChassis);

                // Stages Query
                //string queryStages = $@"SELECT DISTINCT T0.""StageId"", T1.""Desc""
                //               FROM WOR1 T0
                //               INNER JOIN ORST T1 ON T0.""StageId"" = T1.""AbsEntry""
                //               WHERE T0.""DocEntry"" = {productionDocEntry} AND T0.""StageId"" <> '5'
                //               ORDER BY T0.""StageId""";
                string queryStages = $@"SELECT ""StageId"", ""Name"" 
                            FROM WOR4  
                            WHERE ""DocEntry"" = {productionDocEntry} and ""StgEntry"" <> 5 
                            ORDER BY  ""StageId"" ";

                SAPbobsCOM.Recordset rsStages = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsStages.DoQuery(queryStages);

                int rowNo = 1;
                int currentRow = 0;
                while (!rsChas.EoF)
                {
                    string chassisNo = rsChas.Fields.Item("U_ChasisNo").Value.ToString();
                    string engineNo = rsChas.Fields.Item("U_EngineNo").Value.ToString();
                    string lot = rsChas.Fields.Item("U_LotNo").Value.ToString();
                    string enChDocEntry = rsChas.Fields.Item("U_EnChNo").Value.ToString();
                    int workCounter = 1;

                    rsStages.MoveFirst();
                    while (!rsStages.EoF)
                    {
                        oChild = oChildren.Add();

                        string routeId = rsStages.Fields.Item("StageId").Value.ToString();
                        string routeDesc = rsStages.Fields.Item("Name").Value.ToString();

                        // Get Barcodes from OSRT
                        string queryBarcodes = $@"Select ""U_JobStartBarCode"", ""U_JobStopBarCode"", ""U_JobPauseBarCode"", ""U_JobResumeBarCode"" 
                                               from ORST T where T.""AbsEntry"" in (select ""StgEntry"" from ""WOR4""  where ""StageId"" = '{routeId}' and ""DocEntry"" ={productionDocEntry})";
                        SAPbobsCOM.Recordset rsBarcode = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        rsBarcode.DoQuery(queryBarcodes);

                        oChild.SetProperty("U_SrNo", rowNo.ToString());
                        oChild.SetProperty("U_EngineNo", engineNo);
                        oChild.SetProperty("U_ChasisNo", chassisNo);
                        oChild.SetProperty("U_RouteId", routeId + " - " + routeDesc);
                        oChild.SetProperty("U_RouteIdNum", routeId);
                        oChild.SetProperty("U_RouteDsp", routeDesc);
                        oChild.SetProperty("U_BatchNo", lot);
                        oChild.SetProperty("U_Status", "Pending");
                        oChild.SetProperty("U_AdvancedSBNo", enChDocEntry);

                        oChild.SetProperty("U_FirstBarcode", rsBarcode.Fields.Item("U_JobStartBarCode").Value.ToString());
                        oChild.SetProperty("U_LastBarcode", rsBarcode.Fields.Item("U_JobStopBarCode").Value.ToString());
                        oChild.SetProperty("U_JobPause", rsBarcode.Fields.Item("U_JobPauseBarCode").Value.ToString());
                        oChild.SetProperty("U_JobResume", rsBarcode.Fields.Item("U_JobResumeBarCode").Value.ToString());

                        string chasisSuffix = chassisNo.Length >= 8 ? chassisNo.Substring(chassisNo.Length - 8) : chassisNo;
                        string workID = $"{lot}-{chasisSuffix}-{workCounter.ToString("D2")}";
                        oChild.SetProperty("U_WorkId", workID);

                        workCounter++;
                        rowNo++;
                        currentRow++;
                        rsStages.MoveNext();
                    }
                    oChildren.Add();
                    currentRow++;
                    rsChas.MoveNext();
                }

                oParams = oGeneralService.Add(oGeneralData);
                return Convert.ToInt32(oParams.GetProperty("DocEntry"));

            }
            catch (Exception ex)
            {
                //errorMessage = Utilities.Application.Company.GetLastErrorDescription();
                //if (string.IsNullOrEmpty(errorMessage)) errorMessage = ex.Message;
                //return -1;
                string diError = Utilities.Application.Company.GetLastErrorDescription();
                if (string.IsNullOrEmpty(diError))
                {
                    errorMessage = "General Service Error: " + ex.Message;
                }
                else
                {
                    errorMessage = $"UDO DI Error: {diError} (System Msg: {ex.Message})";
                }
                return -1;
            }
        }
        #endregion

        #region Update OCN 
        private void UpdateUdfViaDI(int docEntry, string value)
        {
            SAPbobsCOM.ProductionOrders oProdOrder = (SAPbobsCOM.ProductionOrders)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oProductionOrders);
            try
            {
                if (oProdOrder.GetByKey(docEntry))
                {
                    oProdOrder.UserFields.Fields.Item("U_OCNCode").Value = value;
                    oProdOrder.Update();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oProdOrder);
            }
        }
        #endregion

    }
}