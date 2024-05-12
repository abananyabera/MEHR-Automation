using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using static System.Net.Mime.MediaTypeNames;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.ComponentModel.Design;
using static System.Windows.Forms.LinkLabel;
using System.CodeDom.Compiler;
using System.Diagnostics.Metrics;



namespace MEHR_Automation
{
    class Program
    {

        static void Main(string[] args)
        {
            tablebackup tableBackup = new tablebackup();
            DataLoadCount dataLoadCount = new DataLoadCount();
            VerifyingDuplicateData Datachecking = new VerifyingDuplicateData();
            CountofNewRecords countofNewRecords = new CountofNewRecords();
            ImportChanged ImportChanged = new ImportChanged();
            OrigInternetemailNotchanged origInternetemailNotchanged = new OrigInternetemailNotchanged();
            OrigEpassId origEpassId = new OrigEpassId();
            OrigInternetEmail origInternetEmail = new OrigInternetEmail();
            OrigFirst origFirst = new OrigFirst();
            OrigLast origLast = new OrigLast();
            OrigMiddle origMiddle = new OrigMiddle();
            ImportNotChanged importNotChanged = new ImportNotChanged();
            List_MicroQueries list_MicroQueries = new List_MicroQueries();
            List_pickingQueries list_PickingQueries = new List_pickingQueries();
            VerifyingDuplicateData verifyingDuplicateData = new VerifyingDuplicateData();
            Special_Characters special_Characters = new Special_Characters();


            Console.WriteLine("DataBase Connection is successfull");
            ReadLine();

            Console.WriteLine("****  MEHR DAY 2  ACTIVITY AUTOMATION  ****");
            Console.WriteLine(" \n Next Action : Please click Enter to take the Table backup for the 'tbl_employees_stage1' ");
            ReadLine();



            // Get the user's directory
            string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            StoredProcedure storedProcedure = new StoredProcedure();
            sqlconnection.Open();
            

            tableBackup.TakeTableBackup_tbl_employees_stage1(sqlconnection);//TableBackup1
            Console.WriteLine(" \n Next Action : Please click Enter to take the Table backup for the 'tbl_employees_stage1_hold' ");
            ReadLine();
            tableBackup.TakeTableBackup_tbl_employees_stage1_hold(sqlconnection);//TableBackup2
            Console.WriteLine(" \n Next Action : Please click Enter to compare the Triage file and the Query Result");
            ReadLine();

            bool DataLoadcountcomparision = dataLoadCount.Dataloadfile(sqlconnection);
            if (DataLoadcountcomparision == true)
            {
                Console.WriteLine("\nDataLoad_77_After_triage file count and Query count is Equal.");
                Console.WriteLine(" \n Next Action : Please click Enter to Disable the Jobs ");
                ReadLine();
            }
            else
            {
                Console.WriteLine("\nDataLoad_77_After_triage file count and Query count is not Equal.");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                ReadLine();

                //Environment.Exit(0);
            }
            
            Console.WriteLine("\n NOTE:  JOBS HAS TO BE DISABLED MANUALLY ");

            Console.WriteLine("\n JOBS: \n 1. UserID Synch Job \n 2. Modification Synch Job \n Please press Y to continue");
            char chkJobs = Console.ReadLine()[0];
            if (chkJobs == 'Y' || chkJobs == 'y')
            {
                Console.WriteLine("UserID Synch and Modification Synch Jobs are disabled successfully");
                Console.WriteLine("\n Next Action : Please click Enter to Verify the Duplicates in the table");
                ReadLine();

                Datachecking.VerifyDuplicateDataintable(sqlconnection);
                Console.WriteLine("Next Action : Please click Enter to execute the proc_Pre_Update_Processing and procProcessEmployeeUpdates stored procedures");
                Console.WriteLine(" \n NOTE : PLease Execute the stored procedures only once ");
                ReadLine();


                storedProcedure.StoredProcedureExecution(sqlconnection); // stored procedure Execution of proc_Pre_Update_Processing
                Console.WriteLine("proc_Pre_Update_Processing and procProcessEmployeeUpdates Stored procedure are Executed Only Once successfully ");
                Console.WriteLine("\n Next Action: please click Enter to count the records in the table 'tbl_Employees_Import_Add' ");
                ReadLine();


                countofNewRecords.CountNewRecords(sqlconnection); //count of New Records
                Console.WriteLine("\n Next Action: please click Enter to  Execute the Distinct(fieldname) for the table 'tbl_Employees_Import_Changed' ");
                ReadLine();

                ImportChanged.Import_Changed(sqlconnection);
                Console.WriteLine("\n Next Action: please click Enter to verify the data in the 'tbl_Employees_Import_Changed'");
                ReadLine();

                int count = 0;
                while (count < 5)
                {
                    Console.WriteLine("select the fieldname that need to be executed \n 1.orig_epassid \n 2.orig_internet_email \n 3.orig_first \n 4.orig_last \n 5.orig_middle \n");
                    int select_Import_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_changed)
                    {
                        case 1:
                            Console.WriteLine("Next Action: please click Enter to verify the orig_epassid in the 'tbl_Employees_Import_Changed' ");
                            ReadLine();
                            origEpassId.OrigEpassId_Query(sqlconnection);
                            break;

                        case 2:
                            Console.WriteLine("Next Action: please click Enter to verify the orig_internet_email in the 'tbl_Employees_Import_Changed' ");
                            ReadLine();
                            origInternetEmail.OrigInternetEmail_Query(sqlconnection);
                            break;

                        case 3:
                            Console.WriteLine("Next Action: please click Enter to verify the orig_first in the 'tbl_Employees_Import_Changed' ");
                            ReadLine();
                            origFirst.OrigFirst_Query(sqlconnection);
                            break;

                        case 4:
                            Console.WriteLine("Next Action: please click Enter to verify the orig_last in the 'tbl_Employees_Import_Changed' ");
                            ReadLine();
                            origLast.OrigLast_Query(sqlconnection);
                            break;

                        case 5:
                            Console.WriteLine("Next Action: please click Enter to verify the orig_middle in the 'tbl_Employees_Import_Changed' ");
                            ReadLine();
                            origMiddle.OrigMiddle_Query(sqlconnection);
                            break;

                        default:
                            Environment.Exit(0);
                            break;
                    }
                    count++;
                }

                Console.WriteLine("Next Action: please click Enter to  Execute the distinct(fieldname) of 'tbl_Employees_Import_Not_Changed'");
                ReadLine();
                importNotChanged.Import_Not_Changed(sqlconnection);
                //update
                Console.WriteLine("Next Action: please click Enter to verify the orig_internetemail_import_notchanged in the 'tbl_Employees_Import_Not_Changed'");
                ReadLine();

                int count1 = 0;
                while (count1 < 1)
                {
                    Console.WriteLine("select the fieldname that need to be executed  \n 1.orig_internetemail_import_notchanged ");
                    int select_Import_Not_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_Not_changed)
                    {
                        case 1:
                            origInternetemailNotchanged.Orig_internet_Email_Not_changed(sqlconnection);
                            break;

                        default:
                            Environment.Exit(0);
                            break;
                    }
                    count1++;
                }


                //FIX PROCPROCESS EMPLOYEEUPDATES TO GET RID OF USING THIS PROC
                Console.WriteLine("Next Action: Please click Enter To run the stored procedure procUpdatetempfixStoredProcedure");
                ReadLine();
                storedProcedure.procUpdatetempfixStoredProcedure(sqlconnection);
                ReadLine();
                Console.WriteLine("Next Action: Please click Enter To run the stored procedure procUpdatetempfixStoredProcedure");
                storedProcedure.procPostUpdateStoredProcedure(sqlconnection);
                ReadLine();

                Console.WriteLine("Next Action : Please click Enter to execute the Micro Queries");
                ReadLine();


                //Micro Queries Execution
                list_MicroQueries.List_Micro_Queries(sqlconnection);


                Console.WriteLine("Next Action : Please click Enter to execute the 'procRemoveInvalidEmail' Stored Procedure");
                ReadLine();
                storedProcedure.procRemoveInvalidEmail_SP(sqlconnection);


                Console.WriteLine("Next Action : Please click Enter to execute the 'procRefineManager' Stored Procedure");
                ReadLine();
                storedProcedure.procRefineManager_SP(sqlconnection);


                Console.WriteLine("Next Action : Please click Enter to execute the 'procOverrideBadUpdates' Stored Procedure");
                ReadLine();
                storedProcedure.procOverrideBadUpdates_SP(sqlconnection);

                Console.WriteLine("\n **** Next Action: PLease comlete the 'MANUAL RECONCILE' in the Legal_Eagle database  manually **** \n");
                Console.WriteLine("\n Query for MANUAL RECONCILE: \n \n select ETA.masterid[Master_ID], MAD.lglid[Legal_ID], ETA.Employee [Employee_Name_ME], MAD.Employee [Employee_Name_LE], ETA.Email [Email_ME], MAD.Email [Email_LE], ETA.Entity ,ETA.SBU from dbo.view_Employees_ToAdd ETA inner join dbo.view_Manual_Adds MAD on ((ETA.Employee= MAD.Employee and ETA.Email=MAD.Email) or (ETA.Employee= MAD.Employee))");
                Console.WriteLine("\n In the Legal_Eagle Database if it returns then run the stored procedure 'procEmployee_Reconcile' \n QUERY : exec procEmployee_Reconcile");
                ReadLine();


                Console.WriteLine("Next Action : Please click Enter to execute the 'procStage1ReviewCleanup' Stored Procedure");
                ReadLine();
                storedProcedure.procStage1ReviewCleanup_SP(sqlconnection);
                

                Console.WriteLine("Next Action : Please click Enter to execute the 'procUpdatePicklistValues' Stored Procedure");
                ReadLine();
                storedProcedure.procUpdatePicklistValues_SP(sqlconnection);
 
                //pickling Queries Execution
                list_PickingQueries.List_picking_Queries(sqlconnection);

                //code for Special Characters
                Console.WriteLine("Next Action : Please click Enter to execute the 'replaceSpecialChar' Stored Procedure");
                ReadLine();
                storedProcedure.replaceSpecialChar(sqlconnection);

                Console.WriteLine("Next Action : Please click Enter to execute the 'findSpecialChar' Stored Procedure");
                ReadLine();
                special_Characters.findSpecialChars(sqlconnection);

                Console.WriteLine("Next Action : Please click Enter to execute the 'reverifyDuplicates' Stored Procedure");
                ReadLine();
                verifyingDuplicateData.reverifyDuplicates(sqlconnection);

                //Final Query
                //commenting as of now

                Console.WriteLine("Next Action : Please click Enter to execute the 'procUpdateMasterEmployeeTable' Stored Procedure");
                ReadLine();
                //storedProcedure.procUpdateMasterEmployeeTable(sqlconnection);

                Console.WriteLine("\n ** PLEASE ENABLE THE JOBS UserID Synch Job and Modification Synch Job **");

                Console.WriteLine( " \n ** MEHR AUTOMATION IS COMPLETED ** ");

                sqlconnection.Close();

                Console.ReadLine();
            }

            else
            {
                Console.WriteLine("You have not disabled the job. Please disable the Mentioned jobs to proceed further ");
                Console.WriteLine("please type any key to Exit");
                ReadLine();
                //Environment.Exit(0);

            }

        }
        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }



    }

}