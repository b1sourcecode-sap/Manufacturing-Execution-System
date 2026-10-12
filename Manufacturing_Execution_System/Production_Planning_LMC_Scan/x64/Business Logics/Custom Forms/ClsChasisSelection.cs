using SAPbouiCOM;
using SAPbobsCOM;
using System;
using System.Xml;

namespace Production_Planning_LMC
{
    class ClsChasisSelection : Base
    {
        #region constructor and distructor
        public ClsChasisSelection() : base()
       {

       }
        ~ClsChasisSelection()
        {

        }
        #endregion

        public string ParentItemCode { get; set; }
        public string ParentFormUID { get; set; }
        public int ParentRowIndex { get; set; }

        public double RemainingQty { get; set; }

        public void FormDefault()
        {
            LoadApprovalData();
        }

        #region Load Approval Data
        private void LoadApprovalData()
        {
            //string itemCode = this.Form.DataSources.DBDataSources.Item("OWOR").GetValue("ItemCode", 0).Trim();
            var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtxSel").Specific;

            // Step 1: Prepare DataTable (with safe check)
            SAPbouiCOM.DataTable dt;

            try
            {
                dt = _Form.DataSources.DataTables.Item("dtDisplay");
                dt.Clear();  // reuse if already exists
            }
            catch
            {
                // create if doesn't exist
                dt = _Form.DataSources.DataTables.Add("dtDisplay");
                dt.Columns.Add("U_Selected", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
                dt.Columns.Add("U_ChasisNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_EngineNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_TransNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_SetKey", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                //dt.Columns.Add("U_SerialType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_LotNo", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
            }

            // Step 2: Bind Matrix Columns to DataTable
            matrix.Columns.Item("chkSelect").DataBind.Bind("dtDisplay", "U_Selected");
            matrix.Columns.Item("colChasis").DataBind.Bind("dtDisplay", "U_ChasisNo");
            matrix.Columns.Item("colEngine").DataBind.Bind("dtDisplay", "U_EngineNo");
            matrix.Columns.Item("colTrans").DataBind.Bind("dtDisplay", "U_TransNo");
            matrix.Columns.Item("colSetKey").DataBind.Bind("dtDisplay", "U_SetKey");
            //matrix.Columns.Item("colSerial").DataBind.Bind("dtDisplay", "U_SerialType");
            matrix.Columns.Item("colLotNo").DataBind.Bind("dtDisplay", "U_LotNo");
            Recordset rs = (Recordset)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);
            //added condition where model equals itemcode of the form
            string query = $"SELECT T0.\"U_ChasisNo\", T0.\"U_EngineNo\", T0.\"U_SetKey\", T0.\"U_TransNo\", T1.\"U_LotNo\",T1.\"DocEntry\", T0.\"U_PinCode\" FROM \"@ENGCHASISMMC\" T0 INNER JOIN \"@ENGCHASISMMH\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" WHERE T0.\"U_ModelCode\" = '{ParentItemCode}' and T0.\"U_ProdOrdNo\" IS NULL ";
            rs.DoQuery(query);

            // Step 4: Load Data into DataTable
            while (!rs.EoF)
            {
                int row = dt.Rows.Count;
                dt.Rows.Add();
                dt.SetValue("U_ChasisNo", row, rs.Fields.Item("U_ChasisNo").Value.ToString());
                dt.SetValue("U_EngineNo", row, rs.Fields.Item("U_EngineNo").Value.ToString());
                dt.SetValue("U_TransNo", row, rs.Fields.Item("U_TransNo").Value.ToString());
                dt.SetValue("U_SetKey", row, rs.Fields.Item("U_SetKey").Value.ToString());
                //dt.SetValue("U_SerialType", row, rs.Fields.Item("U_CreatedBy").Value.ToString());
                dt.SetValue("U_LotNo", row, rs.Fields.Item("U_LotNo").Value.ToString());

                rs.MoveNext();
            }

            // Step 5: Load Matrix from DataTable
            matrix.LoadFromDataSource();
        }

        #endregion

        public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            if (pVal.ItemUID == "mtxSel" && pVal.ColUID == "chkSelect" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
            {
                Matrix oMat = (Matrix)Form.Items.Item("mtxSel").Specific;
                SAPbouiCOM.DataTable dt = Form.DataSources.DataTables.Item("dtDisplay");

                int index = pVal.Row - 1;

                if (index < 0 || index >= dt.Rows.Count)
                    return;

                SAPbouiCOM.CheckBox chk =
                    (SAPbouiCOM.CheckBox)oMat.Columns.Item("chkSelect").Cells.Item(pVal.Row).Specific;

                string newVal = chk.Checked ? "Y" : "N";

                dt.SetValue("U_Selected", index, newVal);

                oMat.LoadFromDataSource();
            }

            if (pVal.EventType == BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction)
            {
                if (pVal.ItemUID == "btnSelect" || pVal.ItemUID == "1")
                {
                    TransferSelectedData();
                }
            }



        }

        private void TransferSelectedData()
        {
            SAPbouiCOM.Form parentForm = null;
            try
            {
                // Get DataTable from Child Form
                Matrix childMatrix = (Matrix)this.Form.Items.Item("mtxSel").Specific;
                childMatrix.FlushToDataSource(); // Ensure checkbox changes are committed to DT
                SAPbouiCOM.DataTable dt = this.Form.DataSources.DataTables.Item("dtDisplay");

                // Count Selected Rows
                int selectedCount = 0;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    if (dt.GetValue("U_Selected", i).ToString().Trim() == "Y")
                    {
                        selectedCount++;
                    }
                }

                if (selectedCount == 0)
                {
                    Utilities.Application.SBO_Application.SetStatusBarMessage("No chassis selected.", BoMessageTime.bmt_Short, true);
                    return;
                }

                // Validate against Planned Qty
                if (selectedCount > this.RemainingQty)
                {
                    Utilities.Application.SBO_Application.MessageBox($"Error: You selected {selectedCount} items, but only {this.RemainingQty} quantity is remaining in the Production Order.");
                    return; // Stop processing
                }

                // Get Parent Form and Matrix
                parentForm = Utilities.Application.SBO_Application.Forms.Item(this.ParentFormUID);
                SAPbouiCOM.Matrix parentMatrix = (SAPbouiCOM.Matrix)parentForm.Items.Item("mtxEngChas").Specific;

                parentForm.Freeze(true);

                int currentParentRow = this.ParentRowIndex;
                bool isFirstSelection = true;

                // Loop and Transfer
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string isSelected = dt.GetValue("U_Selected", i).ToString().Trim();

                    if (isSelected == "Y")
                    {
                        if (isFirstSelection)
                        {
                            // Fill the EXACT row where the user pressed Tab
                            FillParentRow(parentMatrix, currentParentRow, dt, i);
                            isFirstSelection = false;
                        }
                        else
                        {
                            // Create NEW rows for additional selections
                            parentMatrix.AddRow();
                            currentParentRow = parentMatrix.RowCount; // Update pointer to the new last row

                            // Auto-number the Serial No
                            ((SAPbouiCOM.EditText)parentMatrix.Columns.Item("colSrNo").Cells.Item(currentParentRow).Specific).Value = currentParentRow.ToString();

                            FillParentRow(parentMatrix, currentParentRow, dt, i);
                        }
                    }
                }

                // Close Child Form
                this.Form.Close();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox("Error transferring data: " + ex.Message);
            }
            finally
            {
                if (parentForm != null) parentForm.Freeze(false);
            }
        }

        private void FillParentRow(SAPbouiCOM.Matrix mtx, int rowIdx, SAPbouiCOM.DataTable dt, int dtRowIdx)
        {
            // Map DataTable columns to Parent Matrix Columns
            try
            {
                string chasis = dt.GetValue("U_ChasisNo", dtRowIdx).ToString();
                string engine = dt.GetValue("U_EngineNo", dtRowIdx).ToString();
                string trans = dt.GetValue("U_TransNo", dtRowIdx).ToString();
                string setKey = dt.GetValue("U_SetKey", dtRowIdx).ToString();
                string lotNo = dt.GetValue("U_LotNo", dtRowIdx).ToString();
                //string docEntry = dt.GetValue("DocEntry", dtRowIdx).ToString(); 
                string pinCode = dt.GetValue("U_PinCode", dtRowIdx).ToString();

                ((SAPbouiCOM.EditText)mtx.Columns.Item("colCha").Cells.Item(rowIdx).Specific).Value = chasis;
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colEng").Cells.Item(rowIdx).Specific).Value = engine;
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colTrans").Cells.Item(rowIdx).Specific).Value = trans;
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colSet").Cells.Item(rowIdx).Specific).Value = setKey;
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colLotNo").Cells.Item(rowIdx).Specific).Value = lotNo;
                //((SAPbouiCOM.EditText)mtx.Columns.Item("colEgChNo").Cells.Item(rowIdx).Specific).Value = docEntry;
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colPinCod").Cells.Item(rowIdx).Specific).Value = pinCode;

                // Set Action/Delete status
                ((SAPbouiCOM.EditText)mtx.Columns.Item("colDel").Cells.Item(rowIdx).Specific).Value = "Delete";
            }
            catch (Exception ex)
            {
                // Handle individual cell errors (e.g. if a column is missing)
            }
        }


        //private void Create_ReceiptFrom_Production(int prodOrderDocEntry)
        //{
        //    try
        //    {
        //        // 1️⃣ Fetch Production Order Header
        //        SAPbobsCOM.Recordset rsOrder = (SAPbobsCOM.Recordset)
        //            Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        rsOrder.DoQuery($@"
        //    SELECT 
        //        T0.""DocEntry"",
        //        T0.""ItemCode"",
        //        T0.""PlannedQty"",
        //        T0.""Warehouse""
        //    FROM OWOR T0
        //    WHERE T0.""DocEntry"" = {prodOrderDocEntry}
        //");

        //        if (rsOrder.EoF)
        //            throw new Exception("Production Order not found.");

        //        string parentItem = rsOrder.Fields.Item("ItemCode").Value.ToString();
        //        double plannedQty = Convert.ToDouble(rsOrder.Fields.Item("PlannedQty").Value);
        //        string whs = rsOrder.Fields.Item("Warehouse").Value.ToString();


        //        // 2️⃣ Create Receipt from Production (Inventory Gen Entry)
        //        SAPbobsCOM.Documents oReceipt = (SAPbobsCOM.Documents)
        //            Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

        //        oReceipt.DocDate = DateTime.Now;
        //        oReceipt.Comments = $"Auto Receipt — Production Completed | ProdOrder: {prodOrderDocEntry}";
        //        oReceipt.UserFields.Fields.Item("U_ProdOrdNo").Value = prodOrderDocEntry.ToString();


        //        // 3️⃣ Add Receipt Line (Finished Good)
        //        oReceipt.Lines.BaseType = 202;            // Production Order
        //        oReceipt.Lines.BaseEntry = prodOrderDocEntry;
        //        oReceipt.Lines.BaseLine = 0;              // Finished good is always line 0

        //        oReceipt.Lines.ItemCode = parentItem;
        //        oReceipt.Lines.WarehouseCode = whs;
        //        oReceipt.Lines.Quantity = plannedQty;


        //        // 4️⃣ If parent item is serial/batch managed → Assign serials
        //        SAPbobsCOM.Recordset rsItem = (SAPbobsCOM.Recordset)
        //            Utilities.Application.Company.GetBusinessObject(BoObjectTypes.BoRecordset);

        //        rsItem.DoQuery($@"
        //    SELECT ""ManSerNum"", ""ManBtchNum""
        //    FROM OITM 
        //    WHERE ""ItemCode"" = '{parentItem}'");

        //        bool isSerial = rsItem.Fields.Item("ManSerNum").Value.ToString() == "Y";
        //        bool isBatch = rsItem.Fields.Item("ManBtchNum").Value.ToString() == "Y";


        //        // =========================
        //        // 🔵 SERIAL MANAGED
        //        // =========================
        //        if (isSerial)
        //        {
        //            for (int i = 0; i < plannedQty; i++)
        //            {
        //                string newSerial = parentItem + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") + "-" + i;

        //                oReceipt.Lines.SerialNumbers.InternalSerialNumber = newSerial;
        //                oReceipt.Lines.SerialNumbers.ManufacturerSerialNumber = newSerial;
        //                oReceipt.Lines.SerialNumbers.Quantity = 1;
        //                oReceipt.Lines.SerialNumbers.Add();
        //            }
        //        }

        //        else if (isBatch)
        //        {
        //            string batch = "BATCH-" + DateTime.Now.ToString("yyyyMMdd");

        //            oReceipt.Lines.BatchNumbers.BatchNumber = batch;
        //            oReceipt.Lines.BatchNumbers.Quantity = plannedQty;
        //            oReceipt.Lines.BatchNumbers.Add();
        //        }

        //        // Add line
        //        oReceipt.Lines.Add();


        //        // 5️⃣ Add Document
        //        int ret = oReceipt.Add();
        //        if (ret != 0)
        //        {
        //            Utilities.Application.Company.GetLastError(out int errCode, out string errMsg);
        //            throw new Exception($"Receipt creation failed: {errMsg}");
        //        }

        //        Utilities.Application.SBO_Application.StatusBar.SetText(
        //            "Receipt from Production created successfully.",
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        //    }
        //    catch (Exception ex)
        //    {
        //        Utilities.Application.SBO_Application.StatusBar.SetText(
        //            ex.Message,
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }
        //}


    }
}