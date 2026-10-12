using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Production_Planning_LMC
{
    class ClsProductionMaster : Base
    {
        #region Constructor and Distructor
        public ClsProductionMaster() : base()
        {

        }

        ~ClsProductionMaster()
        {

        }
        #endregion

        #region VARIABLE DECLARATION
        SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.DataTable dt;
        string engineNo;
        string prodOrdNo;
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            oForm = Utilities.Application.SBO_Application.Forms.ActiveForm;

            try
            {
                oForm.DataSources.DataTables.Add("DT_1");
                oForm.DataSources.UserDataSources.Add("UD_FROM", BoDataType.dt_DATE);
                oForm.DataSources.UserDataSources.Add("UD_TO", BoDataType.dt_DATE);
                oForm.DataSources.UserDataSources.Add("UD_Lot", BoDataType.dt_SHORT_TEXT, 100);
            }
            catch
            {

            }

            dt = oForm.DataSources.DataTables.Item("DT_1");

            // ADD COLUMNS FIRST
            try
            {
                dt.Columns.Add("RowNum", BoFieldsType.ft_AlphaNumeric, 10);
                dt.Columns.Add("U_LotNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ChasisNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_EngineNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_SetKey", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Model", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ModelCode", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ModelName", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ModelYear", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_TMTYPE", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_BatchNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Status", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ProdOrdNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_WrkOrdNo", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_RepProd", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Delivery", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_MProd", BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_SignOff", BoFieldsType.ft_Date);
                dt.Columns.Add("U_TrimDate", BoFieldsType.ft_Date);
                dt.Columns.Add("U_CCode", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_CCod", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_TMNo", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_SOP", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_CC", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_ModelOCN", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_TRIM", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_Seq", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("DispatchTo", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_DName", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("U_DNum", BoFieldsType.ft_AlphaNumeric, 50);
                dt.Columns.Add("DispatchDate", BoFieldsType.ft_AlphaNumeric, 50);

            }
            catch
            {

            }

            SAPbouiCOM.EditText oEditLot = (SAPbouiCOM.EditText)oForm.Items.Item("txtLotNo").Specific;

            // 1. Perform the binding
            oEditLot.DataBind.SetBound(true, "", "UD_Lot");

            // 2. RE-LINK the CFL (This forces the icon to appear)
            oEditLot.ChooseFromListUID = "CFL_0";
            oEditLot.ChooseFromListAlias = "Code";

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("mtx1").Specific;

            BindMatrix();
        }
        #endregion

        #region Item Event
        public override void Item_Event(string FormUID, ref ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.EventType == BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                {
                    SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                    SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                    if (oDataTable != null)
                    {
                        string val = System.Convert.ToString(oDataTable.GetValue("Code", 0));
                        if (pVal.ItemUID == "txtLotNo")
                        {
                            oForm.DataSources.UserDataSources.Item("UD_Lot").ValueEx = val;
                            LoadMatrix(); // Refresh matrix after selection
                        }
                    }
                }

                if (!pVal.BeforeAction &&
                    pVal.EventType == BoEventTypes.et_VALIDATE)
                {
                    if (pVal.ItemUID == "txtLotNo" || pVal.ItemUID == "Item_3" ||
                        pVal.ItemUID == "Item_5" ||
                        pVal.ItemUID == "Item_7")
                    {
                        LoadMatrix();
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox(ex.Message);
            }

            try
            {
                if (pVal.ItemUID == "btnPrint" && pVal.EventType == BoEventTypes.et_ITEM_PRESSED && pVal.BeforeAction)
                {
                    if (oMatrix.VisualRowCount == 0) return;


                    // 1. Create a List to hold all VINs from the Matrix
                    List<string> vinList = new List<string>();

                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                    {
                        engineNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ColEng").Cells.Item(i).Specific).Value;
                        prodOrdNo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ColProdO").Cells.Item(i).Specific).Value;
                        string vin = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("ColVin").Cells.Item(i).Specific).Value;
                        if (!string.IsNullOrEmpty(vin))
                        {
                            vinList.Add(vin);
                        }
                    }

                    // 2. If we found VINs, call the print function ONCE
                    if (vinList.Count > 0)
                    {
                        PrintAllVinsAtOnce(vinList.ToArray());
                    }
                }

            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.MessageBox(ex.Message);
            }

        }
        #endregion

        #region Bind Matrix
        private void BindMatrix()
        {
            oMatrix.Columns.Item("ColRow").DataBind.Bind("DT_1", "RowNum");
            oMatrix.Columns.Item("ColLot").DataBind.Bind("DT_1", "U_LotNo");
            oMatrix.Columns.Item("ColVin").DataBind.Bind("DT_1", "U_ChasisNo");
            oMatrix.Columns.Item("ColEng").DataBind.Bind("DT_1", "U_EngineNo");
            oMatrix.Columns.Item("ColKey").DataBind.Bind("DT_1", "U_SetKey");
            oMatrix.Columns.Item("ColModel").DataBind.Bind("DT_1", "U_Model");
            oMatrix.Columns.Item("ColModelC").DataBind.Bind("DT_1", "U_ModelCode");
            oMatrix.Columns.Item("ColModelN").DataBind.Bind("DT_1", "U_ModelName");
            oMatrix.Columns.Item("ColModelY").DataBind.Bind("DT_1", "U_ModelYear");
            oMatrix.Columns.Item("ColTMTyp").DataBind.Bind("DT_1", "U_TMTYPE");
            oMatrix.Columns.Item("ColBatch").DataBind.Bind("DT_1", "U_BatchNo");
            oMatrix.Columns.Item("ColStatus").DataBind.Bind("DT_1", "U_Status");
            oMatrix.Columns.Item("ColProdO").DataBind.Bind("DT_1", "U_ProdOrdNo");
            oMatrix.Columns.Item("ColWork").DataBind.Bind("DT_1", "U_WrkOrdNo");
            oMatrix.Columns.Item("ColRecP").DataBind.Bind("DT_1", "U_RepProd");
            oMatrix.Columns.Item("ColDel").DataBind.Bind("DT_1", "U_Delivery");
            ((SAPbouiCOM.EditText)oForm.Items.Item("Item_5").Specific).DataBind.SetBound(true, "", "UD_FROM");
            ((SAPbouiCOM.EditText)oForm.Items.Item("Item_7").Specific).DataBind.SetBound(true, "", "UD_TO");
            oMatrix.Columns.Item("ColMProd").DataBind.Bind("DT_1", "U_MProd");
            oMatrix.Columns.Item("ColSign").DataBind.Bind("DT_1", "U_SignOff");
            oMatrix.Columns.Item("ColTInD").DataBind.Bind("DT_1", "U_TrimDate");
            oMatrix.Columns.Item("ColCCode").DataBind.Bind("DT_1", "U_CCode");
            oMatrix.Columns.Item("ColCCod").DataBind.Bind("DT_1", "U_CCod");
            oMatrix.Columns.Item("ColTMNo").DataBind.Bind("DT_1", "U_TMNo");
            oMatrix.Columns.Item("ColSop").DataBind.Bind("DT_1", "U_SOP");
            oMatrix.Columns.Item("ColCC").DataBind.Bind("DT_1", "U_CC");
            oMatrix.Columns.Item("ColMod").DataBind.Bind("DT_1", "U_ModelOCN");
            oMatrix.Columns.Item("ColModT").DataBind.Bind("DT_1", "U_TRIM");
            oMatrix.Columns.Item("ColSeq").DataBind.Bind("DT_1", "U_Seq");
            oMatrix.Columns.Item("ColDisL").DataBind.Bind("DT_1", "DispatchTo");
            oMatrix.Columns.Item("ColDriD").DataBind.Bind("DT_1", "U_DName");
            oMatrix.Columns.Item("ColDisR").DataBind.Bind("DT_1", "U_DNum");
            oMatrix.Columns.Item("ColDisD").DataBind.Bind("DT_1", "DispatchDate");
        }
        #endregion

        #region Load Matrix
        private void LoadMatrix()
        {
            try
            {
                oForm.Freeze(true);

                string lotNo = ((SAPbouiCOM.EditText)oForm.Items.Item("txtLotNo").Specific).Value.Trim();
                string serialNo = ((SAPbouiCOM.EditText)oForm.Items.Item("Item_3").Specific).Value.Trim();
                string fromDate = ((SAPbouiCOM.EditText)oForm.Items.Item("Item_5").Specific).Value.Trim();
                string toDate = ((SAPbouiCOM.EditText)oForm.Items.Item("Item_7").Specific).Value.Trim();

                if (string.IsNullOrEmpty(lotNo) && string.IsNullOrEmpty(serialNo) && string.IsNullOrEmpty(fromDate) && string.IsNullOrEmpty(toDate))
                {
                    dt.Rows.Clear(); // Clear the data table
                    oMatrix.Clear(); // Clear the matrix
                    return;          // Exit the method without running the query
                }

                StringBuilder query = new StringBuilder();

                // Note: Aliases (AS "U_SignOff") must match the DataTable column names exactly
                query.Append(@"
                SELECT 
                    ROW_NUMBER() OVER() AS ""RowNum"",
                    T0.""U_LotNo"", 
                    T1.""U_ChasisNo"",
                    T1.""U_EngineNo"",
                    T1.""U_SetKey"",
                    T0.""U_Model"",
                    T1.""U_ModelCode"",
                    T1.""U_ModelName"",
                    T1.""U_ModelYear"",
                    T3.""U_TMTYPE"",
                    T8.""U_BatchNo"",
                    T1.""U_Status"",
                    T1.""U_ProdOrdNo"",
                    T1.""U_WrkOrdNo"",
                    T1.""U_RepProd"",
                    T1.""U_Delivery"",
                    MONTHNAME(T4.""DocDate"") AS ""U_MProd"", T6.""U_EndDate"" AS ""U_SignOff"",
                    T5.""U_StartDate"" AS ""U_TrimDate"", SUBSTRING(T1.""U_ModelCode"", LENGTH(T1.""U_ModelCode"") - 2, 3) ""U_CCode"",
                    TRIM(SUBSTRING_REGEXPR('([^\\-]+)$' IN T1.""U_ModelName"")) AS ""U_CCod"",
                    T1.""U_TransNo"" ""U_TMNo"",
                    T0.""U_LotNo"" ""U_SOP"",
                    REPLACE(UPPER(T7.""U_Displacement""), 'CC', '') AS ""U_CC"", T3.""U_MODEL"" ""U_ModelOCN"", T3.""U_TRIM"", T8.""U_SrNo"",
                     T9.""PickRmrk"" ""DispatchTo"",T9.""U_DName"", T9.""U_DNum"", T9.""DocDate"" ""DispatchDate""
                FROM ""@ENGCHASISMMH"" T0
                INNER JOIN ""@ENGCHASISMMC"" T1 ON T0.""DocEntry"" = T1.""DocEntry""
                INNER JOIN ""@LOTNOMASTER"" T2 ON T0.""U_LotNo"" = T2.""Code""
                INNER JOIN ""@OCNMASTER"" T3 ON T2.""U_OCN"" = T3.""Code""
                LEFT JOIN ""OIGN"" T4  ON T1.""U_RepProd"" = T4.""DocEntry""
                LEFT JOIN ""@WRKORDRDTLSC"" T5 ON T5.""U_ChasisNo"" = T1.""U_ChasisNo"" AND T5.""U_FirstBarcode"" = 'T1'
                LEFT JOIN
                (
                    SELECT  A.""U_ChasisNo"", A.""U_EndDate""
                    FROM ""@WRKORDRDTLSC"" A
                    WHERE RIGHT(A.""U_WorkId"", 2) =
                    (
                        SELECT MAX(RIGHT(B.""U_WorkId"", 2))
                        FROM ""@WRKORDRDTLSC"" B
                        WHERE B.""U_ChasisNo"" = A.""U_ChasisNo""
                    )
                ) T6
                ON T6.""U_ChasisNo"" = T1.""U_ChasisNo""
                INNER JOIN OITM T7 ON T1.""U_ModelCode"" = T7.""ItemCode""
                INNER JOIN ""@ENGCHASPO"" T8 ON T1.""U_ChasisNo"" = T8.""U_ChasisNo""
                LEFT JOIN ""ODLN"" T9 ON T1.""U_Delivery"" = T9.""DocEntry""
                WHERE 1 = 1
                ");

                if (!string.IsNullOrEmpty(lotNo))
                    query.Append($@" AND T0.""U_LotNo"" = '{lotNo}' ");

                if (!string.IsNullOrEmpty(serialNo))
                {
                    query.Append($@"
                    AND (
                        T1.""U_ChasisNo"" = '{serialNo}'
                        OR T1.""U_EngineNo"" = '{serialNo}'
                        OR T1.""U_TransNo"" = '{serialNo}'
                        OR T1.""U_SetKey"" = '{serialNo}'
                    )
                ");
                }

                if (!string.IsNullOrEmpty(fromDate))
                    query.Append($@" AND T4.""DocDate"" >= '{fromDate}' ");

                if (!string.IsNullOrEmpty(toDate))
                    query.Append($@" AND T4.""DocDate"" <= '{toDate}' ");

                query.Append(@" ORDER BY ""RowNum"" ASC ");

                dt.ExecuteQuery(query.ToString());

                oMatrix.Clear();
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Utilities.Application.SBO_Application.StatusBar.SetText(ex.Message, BoMessageTime.bmt_Long, BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }
        #endregion

        //private void PrintRowToCrystal(string srNo, string vin, string mYear, string eng, string trans, string key, string model, string trim, string tmType, string cCode, string cName, string cc, string po, string lot, string batch, string mProd, string status, string inv, string dDate, string dTo, string driver, string prodOrdNo)

        //{
        //   // ReportDocument objCryRep = new ReportDocument();

        //    string Database = Utilities.Application.Company.CompanyDB;
        //    string currentUser = Utilities.Application.Company.UserName;

        //    CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
        //    CrystalDecisions.Shared.TableLogOnInfo objTabLoginInfo = new CrystalDecisions.Shared.TableLogOnInfo();

        //    ParameterFields paramFields = new ParameterFields();
        //    ParameterField paramField = new ParameterField();

        //    ParameterDiscreteValue discreteVal = new ParameterDiscreteValue();
        //    ParameterRangeValue rangeVal = new ParameterRangeValue();
        //    using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
        //    {
        //        try
        //        {
        //            string reportPath = System.Windows.Forms.Application.StartupPath + "\\Reports\\ProductionTraveller.rpt";
        //            objCryRep.Load(reportPath);

        //            objCryRep.SetParameterValue("CompanyName", Database);
        //            // VEHICLE IDENTIFICATION
        //            objCryRep.SetParameterValue("SrNo", srNo);
        //            objCryRep.SetParameterValue("ChasisNo", vin);
        //            objCryRep.SetParameterValue("ModelYear", mYear);
        //            objCryRep.SetParameterValue("EngineNo", eng);
        //            objCryRep.SetParameterValue("TransNo", trans);
        //            objCryRep.SetParameterValue("KeySet", key);

        //            // MODEL & SPECIFICATION
        //            objCryRep.SetParameterValue("Model", model);
        //            objCryRep.SetParameterValue("Trim", trim);
        //            objCryRep.SetParameterValue("TMType", tmType);
        //            objCryRep.SetParameterValue("ColorCode", cCode);
        //            objCryRep.SetParameterValue("ColorName", cName);
        //            objCryRep.SetParameterValue("CC", cc);

        //            // PRODUCTION DETAILS
        //            objCryRep.SetParameterValue("PONo", po);
        //            objCryRep.SetParameterValue("LotNo", lot);
        //            objCryRep.SetParameterValue("Batch", batch);
        //            objCryRep.SetParameterValue("Seq", ""); // Add variable if you have it
        //            objCryRep.SetParameterValue("MonthOfProd", mProd);
        //            objCryRep.SetParameterValue("Status", status);

        //            // DISPATCH INFORMATION
        //            objCryRep.SetParameterValue("SalesInvNo", inv);
        //            objCryRep.SetParameterValue("DispatchDate", dDate);
        //            objCryRep.SetParameterValue("DispatchTo", dTo);
        //            objCryRep.SetParameterValue("DriverName", driver);
        //            objCryRep.SetParameterValue("DriverNo", ""); // Add if available

        //            objCryRep.SetParameterValue("GeneratedBy", currentUser);

        //            // Action: Print to default printer without showing dialog
        //            objCryRep.PrintToPrinter(1, false, 0, 0);
        //        }
        //        catch (Exception ex)
        //        {
        //            Utilities.WriteLogFile("Crystal Print Error: " + ex.Message);
        //        }
        //        finally
        //        {
        //            objCryRep.Close();
        //            objCryRep.Dispose();
        //        }
        //    }
        //}

        #region After Print button Click 
        private void PrintRowToCrystal(string srNo, string vin, string mYear, string eng, string trans, string key, string model, string trim, string tmType, string cCode, string cName, string cc, string po, string lot, string batch, string mProd, string status, string inv, string dDate, string dTo, string driver, string prodOrdNo, string seq, string driverno)
        {
            using (CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep = new CrystalDecisions.CrystalReports.Engine.ReportDocument())
            {
                try
                {
                    string reportPath = System.Windows.Forms.Application.StartupPath + "\\Reports\\ProductionTraveller.rpt";
                    objCryRep.Load(reportPath);

                    // DATABASE LOGON LOGIC 
                    CrystalDecisions.Shared.ConnectionInfo objConInfo = new CrystalDecisions.Shared.ConnectionInfo();
                    objConInfo.ServerName = Utilities.Application.Company.Server;
                    objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;
                    objConInfo.UserID = "SYSTEM";
                    objConInfo.Password = "P@ssw0rD";

                    // Apply to Main Report
                    foreach (CrystalDecisions.CrystalReports.Engine.Table oTable in objCryRep.Database.Tables)
                    {
                        CrystalDecisions.Shared.TableLogOnInfo oLogonInfo = oTable.LogOnInfo;
                        oLogonInfo.ConnectionInfo = objConInfo;
                        oTable.ApplyLogOnInfo(oLogonInfo);
                    }

                    // Apply to Subreports
                    foreach (CrystalDecisions.CrystalReports.Engine.ReportDocument subReport in objCryRep.Subreports)
                    {
                        foreach (CrystalDecisions.CrystalReports.Engine.Table oTable in subReport.Database.Tables)
                        {
                            CrystalDecisions.Shared.TableLogOnInfo oLogonInfo = oTable.LogOnInfo;
                            oLogonInfo.ConnectionInfo = objConInfo;
                            oTable.ApplyLogOnInfo(oLogonInfo);
                        }
                    }

                    // SET MAIN REPORT PARAMETERS
                    objCryRep.SetParameterValue("CompanyName", Utilities.Application.Company.CompanyDB);
                    objCryRep.SetParameterValue("SrNo", srNo ?? "");
                    objCryRep.SetParameterValue("ChasisNo", vin ?? "");
                    objCryRep.SetParameterValue("ModelYear", mYear ?? "");
                    objCryRep.SetParameterValue("EngineNo", eng ?? "");
                    objCryRep.SetParameterValue("TransNo", trans ?? "");
                    objCryRep.SetParameterValue("KeySet", key ?? "");
                    objCryRep.SetParameterValue("Model", model ?? "");
                    objCryRep.SetParameterValue("Trim", trim ?? "");
                    objCryRep.SetParameterValue("TMType", tmType ?? "");
                    objCryRep.SetParameterValue("ColorCode", cCode ?? "");
                    objCryRep.SetParameterValue("ColorName", cName ?? "");
                    objCryRep.SetParameterValue("CC", cc ?? "");
                    objCryRep.SetParameterValue("PONo", po ?? "");
                    objCryRep.SetParameterValue("LotNo", lot ?? "");
                    objCryRep.SetParameterValue("Batch", batch ?? "");
                    objCryRep.SetParameterValue("Seq", seq);
                    objCryRep.SetParameterValue("MonthOfProd", mProd ?? "");
                    objCryRep.SetParameterValue("Status", status ?? "");
                    objCryRep.SetParameterValue("SalesInvNo", inv ?? "");
                    objCryRep.SetParameterValue("DispatchDate", dDate ?? "");
                    objCryRep.SetParameterValue("DispatchTo", dTo ?? "");
                    objCryRep.SetParameterValue("DriverName", driver ?? "");
                    objCryRep.SetParameterValue("DriverNo", driverno);
                    objCryRep.SetParameterValue("GeneratedBy", Utilities.Application.Company.UserName);

                    // for subreports
                    int docEntryInt = 0;
                    int.TryParse(prodOrdNo, out docEntryInt);

                    // Target Subreport: "Production Timeline"
                    objCryRep.SetParameterValue("DocKey@", docEntryInt, "Production Timeline");
                    objCryRep.SetParameterValue("EngineNo@", eng ?? "", "Production Timeline");

                    // Target Subreport: "RemarksEvents"
                    objCryRep.SetParameterValue("DocKey@", docEntryInt, "RemarksEvents");
                    objCryRep.SetParameterValue("EngineNo@", eng ?? "", "RemarksEvents");

                    // PRINT
                    objCryRep.PrintToPrinter(1, false, 0, 0);
                }
                catch (Exception ex)
                {
                    Utilities.WriteLogFile("Crystal Print Error: " + ex.Message);
                    Utilities.Application.SBO_Application.StatusBar.SetText("Crystal Error: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
                finally
                {
                    objCryRep.Close();
                    objCryRep.Dispose();
                }
            }
        }
        #endregion


        private void PrintAllVinsAtOnce(string[] vinArray)
        {
            using (ReportDocument objCryRep = new ReportDocument())
            {
                try
                {
                    string reportPath = System.Windows.Forms.Application.StartupPath + "\\Reports\\ProductionTraveller.rpt";
                    objCryRep.Load(reportPath);

                    // DATABASE LOGON (Required so the report can fetch its own data)
                    ConnectionInfo objConInfo = new ConnectionInfo();
                    objConInfo.ServerName = Utilities.Application.Company.Server;
                    objConInfo.DatabaseName = Utilities.Application.Company.CompanyDB;
                    objConInfo.UserID = "SYSTEM";
                    objConInfo.Password = "P@ssw0rD";

                    foreach (Table oTable in objCryRep.Database.Tables)
                    {
                        TableLogOnInfo oLogonInfo = oTable.LogOnInfo;
                        oLogonInfo.ConnectionInfo = objConInfo;
                        oTable.ApplyLogOnInfo(oLogonInfo);
                    }

                    // PASS THE ARRAY OF VINS
                    // Because "Allow Multiple Values" is True in the .rpt, 
                    // Crystal accepts this array and filters the SQL query for all of them.
                    objCryRep.SetParameterValue("chasis@", vinArray);

                    // Pass other global parameters
                    objCryRep.SetParameterValue("CompanyName", Utilities.Application.Company.CompanyDB);
                    objCryRep.SetParameterValue("GeneratedBy", Utilities.Application.Company.UserName);


                    int docEntryInt = 0;
                    int.TryParse(prodOrdNo, out docEntryInt);

                    //subreports 
                    //objCryRep.SetParameterValue("DocKey@", docEntryInt, "Production Timeline");
                    //objCryRep.SetParameterValue("EngineNo@", engineNo ?? "", "Production Timeline");

                    // Target Subreport: "RemarksEvents"
                    //objCryRep.SetParameterValue("DocKey@", docEntryInt, "RemarksEvents");
                    //objCryRep.SetParameterValue("EngineNo@", engineNo ?? "", "RemarksEvents");

                    // PRINT
                    // This sends ONE job to the printer/PDF. 
                    // If there are 5 VINs in the array, the PDF will have 5 pages.
                    objCryRep.PrintToPrinter(1, false, 0, 0);
                }
                catch (Exception ex)
                {
                    Utilities.Application.SBO_Application.StatusBar.SetText("Print Error: " + ex.Message);
                }
            }
        }

    }
}
