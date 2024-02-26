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



namespace MEHR_Automation
{
    class Program
    {
        
        static void Main(string[] args)
        {
            OrigEpassId origEpassId = new OrigEpassId();
            OrigInternetEmail origInternetEmail = new OrigInternetEmail();
            OrigFirst origFirst = new OrigFirst();
            OrigLast origLast = new OrigLast();
            OrigMiddle origMiddle = new OrigMiddle();
            Console.WriteLine("****  MEHR DAY 2  ACTIVITY AUTOMATION  ****");
            Console.ReadLine();
            
            // Get the user's directory
            string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            sqlconnection.Open();
            Console.WriteLine("-------  DataBase Connection is successfull  --------");
            Console.ReadLine();
           

            tablebackup tableBackup = new tablebackup();
            tableBackup.takeTableBackup1(sqlconnection);//TableBackup1
            tableBackup.takeTableBackup2(sqlconnection);//TableBackup2


            DataLoadCount dataLoadCount = new DataLoadCount();

            bool DataLoadcountcomparision = dataLoadCount.Dataloadfile(sqlconnection);
            if(DataLoadcountcomparision == true)
            {
                Console.WriteLine("\nDataLoad_77_After_triage file count and Query count is Equal.");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine("\nDataLoad_77_After_triage file count and Query count is not Equal.");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            

            Console.WriteLine("Please Disable the jobs Manually : \n 1. UserID Synch Job \n 2. Modification Synch Job \n Please press Y to continue");
            char chkJobs = Console.ReadLine()[0];
            if (chkJobs == 'Y' || chkJobs == 'y')
            {
                Console.WriteLine("\n UserID Synch and Modification Synch Jobs are disabled successfully");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                VerifyingDuplicateData Datachecking = new VerifyingDuplicateData();
                Datachecking.VerifyDuplicateDataintable(sqlconnection);


                StoredProcedure storedProcedure = new StoredProcedure(); //Stored Procedure Execution
                storedProcedure.StoredProcedureExecution(sqlconnection);
                Console.WriteLine("\n proc_Pre_Update_Processing and procProcessEmployeeUpdates Stored procedure are Executed once successfully ");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                CountofNewRecords countofNewRecords = new CountofNewRecords(); //count of New Records
                countofNewRecords.CountNewRecords(sqlconnection);


                ImportChanged importChanged = new ImportChanged();
                importChanged.Import_Changed(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                int count = 0;
                while (count < 5)
                {
                    Console.WriteLine("select the fieldname that need to be executed \n 1.orig_epassid \n 2.orig_internet_email \n 3.orig_first \n 4.orig_last \n 5.orig_middle \n");
                    int select_Import_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_changed)
                    {
                        case 1:
                            origEpassId.execQuery(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;

                        case 2:
                            origInternetEmail.execQuery(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;

                        case 3:
                            origFirst.execQuery(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;
                        case 4:
                            origLast.execQuery(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;

                        case 5:
                            origMiddle.execQuery(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;

                        default:
                            Environment.Exit(0);
                            break;

                    }
                    count++;
                }


                Console.ReadLine();

                ImportNotChanged importNotChanged = new ImportNotChanged();
                importNotChanged.Import_Not_Changed(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                int count1 = 0;
                while (count1 < 1)
                {
                    Console.WriteLine("select the fieldname that need to be executed  \n 1.orig_internetemail_import_notchanged ");
                    int select_Import_Not_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_Not_changed)
                    {
                        case 1:
                            OrigInternetemailNotchanged origInternetemailNotchanged = new OrigInternetemailNotchanged();
                            origInternetemailNotchanged.Orig_internet_Email_Not_changed(sqlconnection);
                            Console.WriteLine("-------------------------------------------------------------");
                            break;
                            
                        default:
                            Environment.Exit(0);
                            break;


                    }
                    count1++;
                }
                Console.ReadLine();
                Console.ReadLine();

                //FIX PROCPROCESS EMPLOYEEUPDATES TO GET RID OF USING THIS PROC
                //storedProcedure.procUpdatetempfixStoredProcedure(sqlconnection);
                //Console.WriteLine("-------------------------------------------------------------");
                //Console.ReadLine();
                //storedProcedure.procPostUpdateStoredProcedure(sqlconnection);
                //Console.WriteLine("-------------------------------------------------------------");
                //Console.ReadLine();

                //MACROQUERIES EXECUTION
                MacroQueries macroQueries = new MacroQueries();
                macroQueries.Add_Counts_by_Datasource(sqlconnection); 
                macroQueries.Changed_Fields_by_DataSource(sqlconnection);
                macroQueries.Change_NotUpdated_by_DataSource(sqlconnection);
                macroQueries.AddDeleted_by_DataSource(sqlconnection);
                macroQueries.Removed_Countby_DataSource(sqlconnection);
                macroQueries.Check_Email_Types(sqlconnection); 
                macroQueries.Missing_Email_Types(sqlconnection); Console.ReadLine();
                macroQueries.Coastal_Manager_Duplicates(sqlconnection); Console.ReadLine();
                macroQueries.Danisco_in_Workday_Duplicates(sqlconnection); Console.ReadLine();
                macroQueries.Pioneer_D_Group_Match_With_DuPont(sqlconnection); Console.ReadLine();
                macroQueries.Potential_Duplicates_with_potential_match(sqlconnection); Console.ReadLine();
                macroQueries.MyAccessID_Duplicates(sqlconnection); Console.ReadLine();
                macroQueries.Removal_Not_In_Duplicate_Tables(sqlconnection); Console.ReadLine();
                macroQueries.Email_Duplicates(sqlconnection); Console.ReadLine();
                macroQueries.Add_Delete_Expatriates(sqlconnection); Console.ReadLine();
                macroQueries.New_Expatriates(sqlconnection); Console.ReadLine();
                macroQueries.Removed_Expatriates(sqlconnection); Console.ReadLine();
                macroQueries.vw_AddCountByDataSource(sqlconnection); Console.ReadLine();
                macroQueries.vw_RemoveCountByDatasource(sqlconnection); Console.ReadLine();


                storedProcedure.procRemoveInvalidEmailStoredProcedure(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
                storedProcedure.procRefineManagerDataStoredProcedure(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
                storedProcedure.procOverrideBadUpdatesStoredProcedure(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
                storedProcedure.procStage1ReviewCleanupStoredProcedure(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
                storedProcedure.procUpdatePicklistValuesStoredProcedure(sqlconnection);
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                sqlconnection.Close();



                Console.ReadLine();




            }
            else
            {
                Console.WriteLine("You have not disabled the job. Please disable the Mentioned jobs to proceed further ");
                //Console.WriteLine("please type any key to Exit");
                //Console.WriteLine("-------------------------------------------------------------");
                //Console.ReadLine();
                //Environment.Exit(0);

            }



        }


        


    }
}
