using System;
using System.Data.SqlClient;

namespace Production_Planning_LMC
{
    /// <summary>
    /// Summary description for clsStart.
    /// </summary>
    public class Start
    {
        static void Main()

        {
            try
            {
                // Company connection and other initilizations.
                Utilities.Application = EventListener.getEventListener();

                //1. Checking whether DBLogin.INI file exists or not
                //2. Checking whether DBLogin.INI file is having valid login id and password  
                bool oFlag = true;
                //CustomLogin.GetLoginInfo oLogin = new CustomLogin.GetLoginInfo();
                //oLogin.VerifyConnectionInfo(Utilities.Application.Company.Server, Utilities.Application.Company.CompanyDB, ref oFlag);
                //if (oFlag == true)
                if (true)
                {
                    //Getting Login ID and Password from DBLogin.INI
                    //string[] oLoginInfo = oLogin.GetInfo();
                    //Constants.gSERVER = oLoginInfo[0];
                    //Constants.gUSER_ID = oLoginInfo[1];
                    //Constants.gUSER_PASSWORD = oLoginInfo[2];
                    //Constants.gLicenseServer = oLoginInfo[3];
                    
                    //string Msg = string.Empty;
                   // SAPbobsCOM.Company oCompany = Utilities.Application.Company;

                    
                        GC.Collect();
                        // initialize Database 
                        Database.InitializeDatabase();
                        Utilities.LoadMenus(Constants.Menus.MENUS_ADD);
                        Utilities.ShowSucessMessage("Add-On is Connected Successfully");
                        System.Windows.Forms.Application.Run();
                        //initialize CRDisplay 
                       // Utilities.InitializeCRDisplay();  //USE WHILE WOKING WITH CRYSTAL REPORTS 
                        //Utilities.setFilter();
                        //Add Menu's	

                       
                        //Utilities.CreateMenus();
                        //Utilities.CreateConfig();
                        //Utilities.AddFormatedSearches();

                        //Create Folders
                        //Utilities.CreateFolders();					
                        
                    
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
    }
}
