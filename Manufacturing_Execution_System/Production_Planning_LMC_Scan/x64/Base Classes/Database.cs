using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace Production_Planning_LMC
{
	/// <summary>
	/// Summary description for Database.
	/// </summary>
	public sealed class Database
	{
		private Database() 
		{ 
		} 

		#region INITIALISE DATABASE
		public static void InitializeDatabase()
		{
			SAPbouiCOM.ProgressBar oProgressBar = null;
			string[] oChildTables = new string[3];
			string[] oFindColumns = new string[3];
			string[] oColumnName  = new string[2];
			try
			{
				oProgressBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Initializing database...Please wait.", 50, false);
				oProgressBar.Value = 1;

				if( !Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.StartTransaction();


                #region CREATE TABLE

                //CREATE FASHION PARAMETER TABLE
                //CreateTable("B1F_FSPM", "B1SOL Fashion Parameter", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //AddColumn("B1F_FSPM", "B1FCMSEP", "B1SOL Company Separator", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None,SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("B1F_FSPM", "B1FFSNDF", "B1SOL Fashion Default", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #region  LGBOM Header

                //CreateTable("LGBOMH", "LG BOM Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //AddColumn("LGBOMH", "PRNum", "Product No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "PRName", "Product Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "Whse", "WareHouse", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "UOM", "UOM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "RevNum", "Revision No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "SDate", "Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "VDate", "Valid Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "RevRemarks", "RevRemarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "CretedBy", "Created By", 155, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "PriceLNo", "Price List No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMH", "PriceLName", "Price List Name", 155, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region  LGBOM Child

                //CreateTable("LGBOMC", "LG BOM Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("LGBOMC", "Type", "Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "ItemName", "Item Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "Quantity", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "UOM", "UOM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "Whse", "Warehouse", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGBOMC", "IssueMethod", "Issue Method", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region UDF - ORST
                //AddColumn("ORST", "JobStartBarCode", "Job Start BarCode", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ORST", "JobPauseBarCode", "Job Pause BarCode", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ORST", "JobResumeBarCode", "Job Resume BarCode", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ORST", "JobStopBarCode", "Job Stop BarCode", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF - OWOR
                //AddColumn("OWOR", "RevNum", "Revision Num", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "RevRmrks", "Revision Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "PRNum", "Product No", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "InvtTransferReq", "Inventory Transfer Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "InvtTransfer", "Inventory Transfer Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "WorkOrderEntry", "Work Order Details Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "CompletedQty", "Completed Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "RemainingQty", "Remaining Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "OCNCode", "OCN Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OWOR", "LotNo", "Lot Number", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                // UDF OWTQ
                //AddColumn("OWTQ", "ProdEntry", "Production Order Entry", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);


                #region UDF - OITT
                //AddColumn("OITT", "RevNum", "Revision Num", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OITT", "RevRmrks", "Revision Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OITT", "PRNum", "Product No", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF - OUSR

                //AddColumn("OUSR", "FinanceA", "Finance Approver", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OUSR", "BPAuth", "BP Approver", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);


                //AddColumn("OADM", "MESBDTime", "MES Browser Display Time Out", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OUSR", "Station", "Station", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF - OCRD
                //AddColumn("OCRD", "SyncSrc", "Sync Source", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF - OSRN
                //AddColumn("OSRN", "PinCode", "Pin Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OSRN", "EngChasisMMNo", "Engine Chasis No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("OSRN", "SerialType", "Serial Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OSRN", "OCNCode", "OCN Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("OSRN", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF- OOAT
                //AddColumn("OOAT", "LotNo", "Lot Number", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region UDF- OPDN
                // AddColumn("OPDN", "Model", "Model", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                //UDF OITM
                //AddColumn("OITM", "OCN", "OCN Code", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                //LOTNOMASTER 
                //AddColumn("@LOTNOMASTER", "OCN", "OCN Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("@LOTNOMASTER", "BatchNo", "Batch No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                #region  ECO Header
                //CreateTable("LGECOH", "LG ECO Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //AddColumn("LGECOH", "CItmCod", "Current Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "CItmNam", "Current Item Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "CItmGrp", "Current Item Group", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "NItmCod", "New Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "NItmNam", "New Item Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "NItmGrp", "New Item Group", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "OpType", "Operation Type", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                ////AddColumn("LGBOMH", "SDate", "Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "DocNo", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "DocTime", "Document Time", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOH", "CretedBy", "Created By", 155, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region  LGECO Child
                //CreateTable("LGECOC", "LG ECO Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("LGECOC", "Chk", "Check", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "PrdCod", "Product Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "PrdNam", "Product Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "PrdGrp", "Product Group", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "RevNum", "Revision No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "RevRmrks", "Revision Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "RevStatus", "Revision Status", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "LineNo", "Line No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "RouteNo", "Route No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "RouteName", "Route Name", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "ExItmQty", "Existing Item Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("LGECOC", "NewItmQty", "New Item Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                ////AddColumn("LGECOC", "chkBox", "Check Box", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region  Production Planning Header
                //AddColumn("ProductionPH", "Branch", "Branch", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //CreateTable("ProductionPH", "Production Planning Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //AddColumn("ProductionPH", "Type", "Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PRNum", "Product No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PRName", "Product Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PQty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "Whse", "WareHouse", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "UOM", "UOM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PPNum", "Production Planning No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "SDate", "Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "VDate", "Valid Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "Remarks", "Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "CretedBy", "Created By", 155, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PriceLNo", "Price List No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PriceLName", "Price List Name", 155, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "ARemarks", "Approver Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                ////AddColumn("ProductionPH", "ARemarks", "Approver Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPH", "PRODEntry", "Production Entry", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region  Production Planning Child

                //CreateTable("ProductionPC", "Production Planning Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("ProductionPC", "Type", "Type", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "ItemCode", "Item Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "ItemName", "Item Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "Quantity", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "BaseQty", "Base Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "PlannedQty", "Plan Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "UOM", "UOM", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "Whse", "Warehouse", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPC", "IssueMethod", "Issue Method", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region  Production Planning Approval
                //CreateTable("ProductionPA", "Production Planning Approval", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("ProductionPA", "Code", "Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "CreatedBy", "CreatedBy", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "Remarks", "Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region BP Master Header
                //CreateTable("ABPMASTERH", "BP Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                //AddColumn("ABPMASTERH", "BPCode", "BP Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BPName", "BP Name", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BPPAN", "PAN / VAT Number", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BPType", "BP Type", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "CreatedBy", "Created By", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERH", "CreatedOn", "Created On", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERH", "ApprovedBy", "Approved By", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERH", "ApprovedOn", "Approved On", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERH", "ActiveSince", "Active Since", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "DefPriceList", "Default Price List", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "SyncStatus", "Sync Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BPGroup", "BP Group", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "AcctPayable", "Account Payable", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "IsRejected", "Is Rejected", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "Remarks", "Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);

                //General Tab(Procurement Access)
                //AddColumn("ABPMASTERH", "BillToAddr", "Bill To Address", 254, SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "ShipToAddr", "Ship To Address", 254, SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "ContactPerson", "Contact Person", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "ContactPhone", "Contact Phone", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_Phone, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "ContactEmail", "Contact Email", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                ////Finance Tab(Finance Access Only)
                //AddColumn("ABPMASTERH", "CreditLimit", "Credit Limit", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "CommitLimit", "Commitment Limit", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "PayTerms", "Payment Terms", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BankName", "Bank Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BankAccNo", "Bank A/C No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BankIFSC", "Bank IFSC", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "BankBranch", "Bank Branch", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                //AddColumn("ABPMASTERH", "Focus", "Focus", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                ////Attachment Tab(Procurement Access)
                //CreateTable("ABPMASTERC", "BP Master Child", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
                //AddColumn("ABPMASTERC", "AttachRemarks", "Attachment Remarks", 254, SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERC", "AttachFileType", "Attachment File Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERC", "AttachFilePath", "Attachment File Path", 254, SAPbobsCOM.BoFieldTypes.db_Memo, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERC", "AttachEntry", "Attach Entry", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERC", "FileName", "File Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ABPMASTERC", "AttachDate", "Attachment Date", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region  Production Planning Approval
                //CreateTable("ProductionPA", "Production Planning Approval", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("ProductionPA", "Code", "Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "CreatedBy", "CreatedBy", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ProductionPA", "Remarks", "Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Engine Chasis Mapping Master Header
                //CreateTable("ENGCHASISMMH", "Engine Chasis MM Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //AddColumn("ENGCHASISMMH", "LotNo", "Lot No", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "DocNo", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "Status", "Status", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "PostDate", "Posting Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "ReqQty", "Required Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "CreatedQty", "Created Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "RemQty", "Remaining Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "Focus", "Focus", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "Model", "Model", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "ToCreate", "To Create", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //AddColumn("ENGCHASISMMH", "GRPO", "GRPO NOs", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "NoChasis", "No of Chasis found", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMH", "BatchNo", "Batch Number", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Engine Chasis Mapping Master Child
                //CreateTable("ENGCHASISMMC", "Engine Chasis MM Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("ENGCHASISMMC", "SrNo", "Serial No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "EngineNo", "Engine No", 36, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "ChasisNo", "Chasis No", 36, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "TransNo", "Transmission No", 36, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "SetKey", "Set Key", 36, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "ModelCode", "Model Code", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "Status", "Status", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "PushInvt", "Push to Inventory", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "RepProd", "Receipt from Production", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //AddColumn("ENGCHASISMMC", "ProdOrdNo", "Production Order No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "ModelName", "Model Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "Delivery", "Delivery", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC", "PinCode", "Generated PIN Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC", "ModelYear", "Model Year", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC", "WrkOrdNo", "WOrk Order Details No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Pin Code Generation
                //CreateTable("ENGCHASISMMC1", "VIN Pincode Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //// VIN and Last 6
                //AddColumn("@ENGCHASISMMC1", "VIN", "VIN Number", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "Last6", "Last 6 Digits", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //// Digits 1–3
                //AddColumn("@ENGCHASISMMC1", "D1", "Digit 1", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "D2", "Digit 2", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "D3", "Digit 3", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //// Digits 4–6
                //AddColumn("@ENGCHASISMMC1", "D4", "Digit 4", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "D5", "Digit 5", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "D6", "Digit 6", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //// Calculations: WSum, Mult, AddVal
                //AddColumn("@ENGCHASISMMC1", "WSum", "Weighted Sum", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "Mult", "Multiplier", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "AddVal", "Addition Value", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //// Calculations: ModVal, PinCode
                //AddColumn("@ENGCHASISMMC1", "ModVal", "Modulus Value", 20, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "PinCode", "Generated PIN Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("@ENGCHASISMMC1", "CalcPrime", "Calculate Prime", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASISMMC1", "SerialNo", "Serial No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Work Order Details Header
                //CreateTable("WRKORDRDTLSH", "Work Order Details Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                //AddColumn("WRKORDRDTLSH", "LotNo", "Lot No", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "ProdOrdNo", "Production Order No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "ProductCode", "Product Code", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "ProductName", "Product Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "AdvancedSBNo", "Serial/Batch Creation No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "DocNo", "Document No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "PostDate", "Posting Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "LastWorkId", "Last Work ID", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "LastUpdated", "Last Updated On", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "UpdatedBy", "Updated By", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "Fetch", "Fetch Details", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "PrdOrdEnt", "Production Order Entry", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "PlannedQty", "Planned Qty", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "Model", "Model", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSH", "InvtTransferNo", "Inventory Transfer Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region Work Order Details Child
                //CreateTable("WRKORDRDTLSC", "Work Order Details Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("WRKORDRDTLSC", "SrNo", "Serial No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "EngineNo", "Engine No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "ChasisNo", "Chasis No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "BatchNo", "Batch No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "WorkId", "Work ID", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "RouteId", "Route ID", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "RouteDsp", "Route Description", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "FirstBarcode", "First Part Barcode", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "LastBarcode", "Last Part Barcode", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "JobExeNo", "Job Exe Doc No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "StartDate", "Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "StartTime", "Start Time", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "EndDate", "End Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "EndTime", "End Time", 20, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "TotalBrkDwnT", "Total Break Down Time", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "ActualTimeCon", "Actual Time Consumed", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "AdvancedSBNo", "Serial/Batch Creation No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "JobPause", "Job Pause BarCode", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "JobResume", "Job Resume Barcode", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //AddColumn("WRKORDRDTLSC", "IssueForProdEntry", "Issue For Production Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WRKORDRDTLSC", "RouteIdNum", "Route ID Num", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region ENGINE CHASIS IN PRODUCTION ORDERs
                //CreateTable("ENGCHASPO", "Engine Chasis Selection Form", SAPbobsCOM.BoUTBTableType.bott_NoObject);
                //AddColumn("ENGCHASPO", "SrNo", "Serial No", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "EngineNo", "Engine No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "ChasisNo", "Chasis No", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "KeySet", "Key Set", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "TransNo", "Transmission No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "ProdDocEntry", "Prod Order DocEntry", 10, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "LotNo", "Lot No", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "EnChNo", "Engine Chasis No", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "PinCode", "Pin Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "Status", "Status", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "PushToInvt", "Push To Inventory", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "ReceiptFrmProd", "Receipt from Production", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "BatchNo", "Batch Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "TotalTime", "Total Time Consumed", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("ENGCHASPO", "ProdSeriesNo", "Production Series Number", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region Work Order Job Execution Header UDO

                //CreateTable("WJOBEXEH", "Work Order Job Header", SAPbobsCOM.BoUTBTableType.bott_Document);

                // General Job Details
                //AddColumn("WJOBEXEH", "JobID", "Job ID", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "JobDesc", "Job Description", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                // Production + Work Order + FG Details
                //AddColumn("WJOBEXEH", "ProdOrdNo", "Production Order No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "WorkOrdNo", "Work Order No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "FGCode", "FG Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "FGDesc", "FG Description", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "EngineNo", "Engine No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "ChassisNo", "Chassis No", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "BatchNo", "Batch No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "Qty", "Qty", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //// Status + Dates
                //AddColumn("WJOBEXEH", "DocNo", "Document No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "Status", "Status", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "StartDate", "Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "StartTime", "Start Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "EndDate", "End Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "EndTime", "End Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //// Calculated Time
                //AddColumn("WJOBEXEH", "TotBrkTime", "Total Break Down Time", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "TotActTime", "Total Actual Time Consumed", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                //// Operator
                //AddColumn("WJOBEXEH", "Operator", "Operator", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "IForProdEntry", "Issue for Production Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                ////alphanumeric field for total time consumed
                //AddColumn("WJOBEXEH", "TotTime", "Total Time", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "Remarks", "Remarks", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEH", "RemarksOut", "Out Remarks", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region Work Order Job Execution Child UDO
                //CreateTable("WJOBEXEC", "Work Order Job Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                //AddColumn("WJOBEXEC", "BrkStartDt", "Break Start Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BrkStartTm", "Break Start Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BrkEndDt", "Break End Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BrkEndTm", "Break End Time", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_Time, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BrkTime", "Break Down Time", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BrkRemarks", "Break Down Remarks", 200, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("WJOBEXEC", "BreakCategory", "Break down Category", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Break Down Category
                //CreateTable("BrkCategory", "Break Down Category", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                #endregion

                #region CREATE UDO

                #region LG BOM
                //oChildTables[0] = "LGBOMC";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_PRNum";
                //CreateUDO("LGBOMH", "LG BOM", "LGBOMH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);




                #endregion

                #region Production Planning
                //oChildTables[0] = "ProductionPC";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_PPNum";
                //CreateUDO("ProductionP", "Production Planning", "ProductionPH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);
                #endregion

                #region Approved BP Creation
                //oChildTables[0] = "ABPMASTERC";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_BPCode";
                //CreateUDO("APPROVEDBPC", "Approved BP Creation", "ABPMASTERH", SAPbobsCOM.BoUDOObjType.boud_MasterData, oFindColumns, oChildTables);
                #endregion

                #region Engine Chasis Mapping Master 
                //oChildTables[0] = "ENGCHASISMMC";
                //oChildTables[1] = "ENGCHASISMMC1";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_LotNo";
                //CreateUDO("ENGCHASISMM", "Engine Chasis MM", "ENGCHASISMMH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);
                #endregion

                #region Work Order Details
                //oChildTables[0] = "WRKORDRDTLSC";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_LotNo";
                //CreateUDO("WRKORDRDTLS", "Work Order Details", "WRKORDRDTLSH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);
                #endregion



                #region LG ECO
                //oChildTables[0] = "LGECOC";
                //oFindColumns[0] = "DocEntry";
                ////oFindColumns[1] = "U_DocNum";
                //CreateUDO("LGECOH", "LG ECO", "LGECOH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);
                #endregion

                #region UPDATEUDFS
                // UpdateColumn("B1F_FSPM", "B1FSBRNDA", "SO Brands Activated?", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);

                #endregion

                #region Work Order Job Execution UDO
                //oChildTables[0] = "WJOBEXEC";
                //oFindColumns[0] = "DocEntry";
                //oFindColumns[1] = "U_JobID";
                //CreateUDO("WJOBEXE", "Work Order Job Execution", "WJOBEXEH", SAPbobsCOM.BoUDOObjType.boud_Document, oFindColumns, oChildTables);
                #endregion

                #endregion

                // Utilities.CreateFunction_RemoveSpecialChar();

                if ( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
 
				// Initialize Connection String
                //Constants.gobjSQLCon = new SqlConnection("user id=" + Constants.gUSER_ID + ";data source=" + Constants.gSERVER + ";pwd=" + Constants.gUSER_PASSWORD + ";initial catalog=" + Utilities.Application.Company.CompanyDB);
                //Constants.gConnStr = "user id=" + Constants.gUSER_ID + ";data source=" + Constants.gSERVER + ";pwd=" + Constants.gUSER_PASSWORD + ";initial catalog=" + Utilities.Application.Company.CompanyDB;
				oProgressBar.Stop();
			}
			catch(Exception ex)
			{
				oProgressBar.Stop();

				if( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
				oProgressBar = null;
				GC.Collect();
			}
		}


        #endregion

               

        #endregion
        #region DATABASE CREATION FUNCTIONS

        #region CREATE USER DEFINED TABLE

        private static void CreateTable(string oTable, string oDescription, SAPbobsCOM.BoUTBTableType oType)
		{
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTable = null;

			try
			{
				oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
				
				if( !oUserTable.GetByKey(oTable) )
				{
					oUserTable.TableName        = oTable;
					oUserTable.TableDescription = oDescription;
					oUserTable.TableType        = oType;

					if( oUserTable.Add() != 0 )	
                    
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
																
				}
			}
			catch(Exception ex)
			{
                throw ex;
			}			
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
				oUserTable = null;
				GC.Collect();
			}
		}

        private static void UpdateTable(string oTable, SAPbobsCOM.BoUTBTableType oType)
        {
            SAPbobsCOM.UserTablesMD oUserTable = null;

            try
            {
                oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUserTable.GetByKey(oTable) != null)
                {
                    if (oUserTable.TableType != SAPbobsCOM.BoUTBTableType.bott_MasterData)
                    {
                        //oUserTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
                        if (oUserTable.Remove() != 0)

                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }
        }
		#endregion

		#region ADD COLUMN
		internal  static void AddColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserFieldsMD oUserField = null;
            int Fieldid;
			try
			{

                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tNO)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    oUserField.TableName = oTable;
                    oUserField.Name = oName;
                    oUserField.Description = oDescription;
                    oUserField.Type = oType;
                    oUserField.SubType = oSubType;

                    if (oType != SAPbobsCOM.BoFieldTypes.db_Float)
                        oUserField.EditSize = oSize;
                    else
                        oUserField.EditSize = 0;

                    
                    oUserField.Mandatory = oMandatory;

                    //if (oTable == "OUSR" && oName == "StatusAuth")
                    //{ 
                    //    oUserField.ValidValues.Value = "NO";
                    //    oUserField.ValidValues.Description = "NO";
                    //    oUserField.ValidValues.Add();
                    //    oUserField.ValidValues.Value = "YES";
                    //    oUserField.ValidValues.Description = "YES";
                    //    oUserField.ValidValues.Add();
                    //    oUserField.DefaultValue = "NO";

                    //}
                  
                    if (oUserField.Add() != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }
                
			}
			catch( Exception ex)
			{
                throw ex;
			}
			finally
			{
                if (oUserField!=null)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
				oUserField = null;
				GC.Collect();
			}
		}

        internal static void UpdateColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserFieldsMD oUserField = null;
            string oTable1=oTable ;
            int Fieldid=0;
            try
            {
                
                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    if(oTable!="OCRD") 
                    oTable1 = @"@" + oTable;

                    oUserField.GetByKey(oTable1, Fieldid);
                    if (oTable == "B1F_GNDR" && (oName == "B1FINACT"))
                    {
                        if (oUserField.DefaultValue != "N")
                        {
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }
                    }
                    else if ((oTable == "OCRD" && oName == "Season") || (oTable == "B1F_ITEMSEAS" && oName == "B1FSEASD") || (oTable == "B1F_ITEMCAPS" && (oName == "B1FCAPSD" ||oName =="B1FDTST" || oName == "B1FDTCP")))
                    {
                        if (oUserField.Remove()  != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());

                    }
                    else if (oTable == "OCRD" && (oName == "B1FBUYINP" || oName == "B1FBUYINPN"))
                    {
                        if (oUserField.ValidValues.Value == "")
                        {
                            oUserField.ValidValues.Value = "N";
                            oUserField.ValidValues.Description = "No";
                            oUserField.ValidValues.Add();
                            oUserField.ValidValues.Value = "Y";
                            oUserField.ValidValues.Description = "Yes";
                            oUserField.ValidValues.Add();
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                    else
                    {
                        if((oTable == "B1F_FSPM" && oName=="B1FSKUNUM") || (oTable == "B1F_FSPM" && oName=="B1FATLEN")|| (oTable == "B1F_FSAC" && oName=="B1FSEQ")||(oTable == "B1F_FSIM" && oName=="B1FSKULN")||(oTable == "B1F_VAT1" && oName=="B1FGRPSQ")||(oTable == "B1F_VAT2" && oName=="B1FGRPSQ"))
                        {
                            if (oUserField.EditSize  < oSize)
                            {
                                oUserField.EditSize = oSize;
                                if (oUserField.Update() != 0)
                                    throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                            }
                        }
                        if (oUserField.Description != oDescription)
                        {
                            oUserField.Description = oDescription;
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                }
              

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oUserField != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();
            }
        }

        #endregion

        //#region IS COLUMN EXISTS
        //private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable,out int FieldId)
        //{
        //    SAPbobsCOM.Recordset oRSColumn = null;
        //    string               oSQL      = string.Empty;
        //    FieldId = 0;
        //    try
        //    {
        //        if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
        //            oTable = @"@" + oTable;

        //        oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
        //        Utilities.ExecuteSQL(ref oRSColumn, oSQL);

        //        if ((int)oRSColumn.Fields.Item(0).Value == 0)
        //            return SAPbobsCOM.BoYesNoEnum.tNO;

        //        else
        //        {
        //            FieldId = (int)oRSColumn.Fields.Item(1).Value;
        //            return SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
        //        oRSColumn = null;
        //        GC.Collect();
        //    }
        //}
        //#endregion

        #region IS COLUMN EXISTS
        private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable, out int FieldId)
        {
            SAPbobsCOM.Recordset oRSColumn = null;
            string oSQL = string.Empty;
            FieldId = 0;
            try
            {
                if (oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO)
                    oTable = @"@" + oTable;

                //oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
                oSQL = @"Select " +
                        @"Count(*) "
                        + " , "
                        + @"""FieldID"""
                        + "From CUFD Where "
                        + @"""TableID"""
                         + "="
                        + "'" + oTable.Trim() + "'"
                        + " And "
                        + @"""AliasID"""
                        + "="
                        + "'" + oColumn.Trim() + "'"
                        + " group by "
                        + @"""FieldID""".Trim();


                Utilities.ExecuteSQL(ref oRSColumn, oSQL);

                if ((int)oRSColumn.Fields.Item(0).Value == 0)
                    return SAPbobsCOM.BoYesNoEnum.tNO;

                else
                {
                    FieldId = (int)oRSColumn.Fields.Item(1).Value;
                    return SAPbobsCOM.BoYesNoEnum.tYES;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
                oRSColumn = null;
                GC.Collect();
            }
        }
        #endregion


		#region CREATE UDO
		private static void CreateUDO(string oUniqueID, string oDescription, string oTable, SAPbobsCOM.BoUDOObjType oType, string[] oFindColumns, string[] oChildTables )
		{
			SAPbobsCOM.UserObjectsMD oUserObject = null;
			int                      oCount;

			try
			{
				oUserObject = (SAPbobsCOM.UserObjectsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

				if( !oUserObject.GetByKey(oUniqueID) )
				{
					oUserObject.Code        = oUniqueID;
					oUserObject.Name        = oDescription;
					oUserObject.ObjectType  = oType;
					oUserObject.TableName   = oTable;

					oUserObject.ManageSeries         = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanCancel            = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanClose             = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanDelete            = SAPbobsCOM.BoYesNoEnum.tNO;

					if( oFindColumns != null )
					{
						if( oFindColumns.Length > 0 )
						{
							oUserObject.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

							for(oCount = 0; oCount <= oFindColumns.Length-1; oCount++)
							{
								if( oFindColumns[oCount] != null )
								{
									oUserObject.FindColumns.ColumnAlias = oFindColumns[oCount];
									oUserObject.FindColumns.Add();
								}

								oFindColumns[oCount] = null;
							}
						}
					}

					if( oChildTables != null )
					{
						if( oChildTables.Length > 0 )
						{
							for(oCount = 0; oCount <= oChildTables.Length-1; oCount++)
							{
								if( oChildTables[oCount] != null )
								{
									oUserObject.ChildTables.TableName = oChildTables[oCount];

									if( oCount != oChildTables.Length - 1 )
										oUserObject.ChildTables.Add();
								}

								oChildTables[oCount] = null;
							}
						}
					}

					if( oUserObject.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObject);
				oUserObject = null;
				GC.Collect();
			}
		}
		#endregion

        public static void CreateUserObject(string CodeID, string Name, string TableName, string Child) //used for registration of user defined table
        {
           
            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            try
            {
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

                if (oUserObjectMD.GetByKey(CodeID) == true)
                    return; //vijay 170708

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;


                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = TableName;

                if (Child != "")
                {
                    oUserObjectMD.ChildTables.TableName = Child;
                    //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { }
                    else
                    { Utilities.Application.Company.GetLastError(out lRetCode, out sErrMsg); }
                else
                { }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

		#region IS KEY EXISTS
		private static SAPbobsCOM.BoYesNoEnum KeyExists(string oTable,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.Recordset oRSColumn = null;
			string               oSQL      = string.Empty;
			
			try
			{		
				if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
					oTable = @"@" + oTable;
				oSQL = "Select Count(*) From OUKD Where TableName = '" + oTable + "' AND KeyName = '" + oKeyName + "' ";
				Utilities.ExecuteSQL(ref oRSColumn, oSQL);
            
				if( (int)oRSColumn.Fields.Item(0).Value == 0 )
					return SAPbobsCOM.BoYesNoEnum.tNO;

				else
					return SAPbobsCOM.BoYesNoEnum.tYES;
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
				oRSColumn = null;
				GC.Collect();
			}
		}
		#endregion

		#region ADD KEYS
		private static void AddKey(string oTable, string[] oColumnName,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserKeysMD oUserKey = null;
			int                   oCount;
			try
			{				
				oUserKey = (SAPbobsCOM.UserKeysMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
				if(KeyExists(oTable,oKeyName,oIsSystemTable) == SAPbobsCOM.BoYesNoEnum.tNO)
				{
					oUserKey.TableName = oTable ;
					oUserKey.KeyName   = oKeyName ;
					if( oColumnName != null )
					{
						if( oColumnName.Length > 0 )
						{						
							for(oCount = 0; oCount <= oColumnName.Length-1; oCount++)
							{
								if( oColumnName[oCount] != null )
								{
									oUserKey.Elements.ColumnAlias = oColumnName[oCount];
									oUserKey.Elements.Add();
								}

								oColumnName[oCount] = null;
							}
						}
					}							
					oUserKey.Unique = SAPbobsCOM.BoYesNoEnum.tYES;					
					if( oUserKey.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex ;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKey);
				oUserKey = null;
				GC.Collect();
			}
		}
		#endregion

		#endregion

	}
}
