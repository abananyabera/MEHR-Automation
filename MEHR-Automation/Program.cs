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



namespace MEHR_Automation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Get the user's directory
            string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            sqlconnection.Open();
            Console.WriteLine("Connection is successfull");
            //execute commands
            Console.WriteLine("--------------------------------------------------------------------------------------------");


            tablebackup tableBackup = new tablebackup();
            tableBackup.takeTableBackup1(sqlconnection);//TableBackup1
            tableBackup.takeTableBackup2(sqlconnection);//TableBackup2


            DataLoadCount dataLoadCount = new DataLoadCount();

            bool DataLoadcountcomparision = dataLoadCount.Dataloadfile(sqlconnection);
            if(DataLoadcountcomparision == true)
            {
                Console.WriteLine("DataLoad_77_After_triage file count and Query count is Equal.");
            }
            else
            {
                Console.WriteLine("DataLoad_77_After_triage file count and Query count is not Equal.");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.ReadLine();

                //Environment.Exit(0);
            }


            Console.WriteLine("Please Disable the jobs: \n 1. UserID Synch Job \n 2. Modification Synch Job \n Please press Y to continue");
            char chkJobs = Console.ReadLine()[0];
            if (chkJobs == 'Y' || chkJobs == 'y')
            {
                Console.WriteLine("\n\n\nJob disabled successfully\n\n\n");

                VerifyingDuplicateData Datachecking = new VerifyingDuplicateData();
                Datachecking.VerifyDuplicateDataintable(sqlconnection);


                StoredProcedure storedProcedure = new StoredProcedure(); //Stored Procedure Execution
                storedProcedure.StoredProcedureExecution(sqlconnection);
                Console.WriteLine(" Both the Stored procedure Executed once successfully ");

                CountofNewRecords countofNewRecords = new CountofNewRecords(); //count of New Records
                countofNewRecords.CountNewRecords(sqlconnection);

                

                #region MyRegion
                Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Changed is started \n");
                string Query13 = "Select distinct(fieldname) from tbl_Employees_Import_Changed";
                SqlDataReader datareader13 = ExecuteQuery(Query13, sqlconnection);
                while (datareader13.Read())
                {
                    string fieldValue = datareader13["fieldname"].ToString();
                    Console.WriteLine(fieldValue);

                }
                Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Changed is completed\n");
                #endregion



                int count = 0;
                while (count < 5)
                {
                    Console.WriteLine("select the fieldname that need to be executed \n 1.orig_epassid \n 2.orig_internet_email \n 3.orig_first \n 4.orig_last \n 5.orig_middle \n");
                    int admin_student = int.Parse(Console.ReadLine());
                    switch (admin_student)
                    {
                        case 1:
                            Console.WriteLine("\n orig_epassid is started \n");
                            string Query14 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from dbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid where a.fieldname = 'orig_epassid'";
                            SqlDataReader datareader14 = ExecuteQuery(Query14, sqlconnection);

                            //create excel workbook
                            var excelApp = new Microsoft.Office.Interop.Excel.Application();
                            var workbook = excelApp.Workbooks.Add();
                            var worksheet = (Worksheet)workbook.Sheets[1];

                            //Add column headers
                            for (int i = 0; i < datareader14.FieldCount; i++)
                            {
                                worksheet.Cells[1, i + 1] = datareader14.GetName(i);
                            }

                            // Add data to Excel worksheet
                            int row = 2;
                            while (datareader14.Read())
                            {
                                for (int i = 0; i < datareader14.FieldCount; i++)
                                {
                                    worksheet.Cells[row, i + 1] = datareader14[i];
                                }
                                row++;
                            }

                            // Save Excel workbook
                            string Pathname = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
                            workbook.SaveAs(Pathname);
                            workbook.Close();
                            excelApp.Quit();

                            Console.WriteLine($"Excel file created at: {Pathname}");



                            Console.WriteLine("\n orig_epassid is completed \n");
                            break;
                        case 2:
                            Console.WriteLine("\n orig_internet_email is started \n");
                            string Query15 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_internet_email'";
                            SqlDataReader datareader15 = ExecuteQuery(Query15, sqlconnection);

                            string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp = new Microsoft.Office.Interop.Excel.Application();
                            existingApp.Visible = true;
                            var existingWorkbook = existingApp.Workbooks.Open(existingPath);

                            // Get or create Sheet2
                            Worksheet sheet2;
                            try
                            {
                                // Try to get Sheet2 by index
                                sheet2 = (Worksheet)existingWorkbook.Sheets[2];
                            }
                            catch
                            {
                                // If Sheet2 doesn't exist, add it
                                sheet2 = (Worksheet)existingWorkbook.Sheets.Add(After: existingWorkbook.Sheets[existingWorkbook.Sheets.Count]);
                                sheet2.Name = "Sheet2";
                            }

                            // Add column headers
                            for (int i = 0; i < datareader15.FieldCount; i++)
                            {
                                sheet2.Cells[1, i + 1] = datareader15.GetName(i);
                            }


                            // Add data to Sheet2
                            int row2 = 2;
                            while (datareader15.Read())
                            {
                                for (int i = 0; i < datareader15.FieldCount; i++)
                                {
                                    sheet2.Cells[row2, i + 1] = datareader15[i];
                                }
                                row2++;
                            }

                            // Save the existing Excel workbook
                            existingWorkbook.Save();
                            existingWorkbook.Close();
                            existingApp.Quit();

                            Console.WriteLine($"Excel file updated at: {existingPath}");
                            Console.WriteLine("\n orig_internet_email is completed \n");
                            break;
                        case 3:
                            Console.WriteLine("\n orig_first is started \n");
                            string Query16 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_first'";
                            SqlDataReader datareader16 = ExecuteQuery(Query16, sqlconnection);
                            while (datareader16.Read())
                            {
                                Console.WriteLine("Tharun");
                            }
                            Console.WriteLine("\n orig_first is completed \n");
                            break;
                        case 4:
                            Console.WriteLine("\n orig_last is started \n");
                            string Query17 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_last'";
                            SqlDataReader datareader17 = ExecuteQuery(Query17, sqlconnection);
                            while (datareader17.Read())
                            {
                                Console.WriteLine("Kanna Sai");
                            }
                            Console.WriteLine("\n orig_last is completed \n");
                            break;
                        case 5:
                            break;
                        default:
                            Environment.Exit(0);
                            break;

                    }
                    count++;
                }



                sqlconnection.Close();


                //Read Excel  people Report File
                //ReadExcelFile();

                Console.ReadLine();




            }
            else
            {
                Console.WriteLine("You have not disabled the job. Please disable the Mentioned jobs to proceed further ");
                //Console.WriteLine("please type any key to Exit");
                //Console.ReadLine();
                //Environment.Exit(0);

            }



        }

        public static SqlDataReader ExecuteQuery(string query, SqlConnection connection) {
            try
            {
                SqlCommand cmd = new SqlCommand(query, connection);
                SqlDataReader datareader = cmd.ExecuteReader();
                return datareader;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                throw;
            }

        }

        //public static void ReadExcelFile()
        //{
        //    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
        //    app.Visible = true;

        //    string path = "C:\\Users\\kwr579\\Desktop\\AUTOMATION\\People Report_1215.xlsx";
        //    Workbook wb;
        //    Worksheet ws;

        //    try
        //    {
        //        wb = app.Workbooks.Open(path);
        //        ws = wb.Worksheets["sheet1"];

        //        string cellData = " " + ws.Range["A1"].Value;
        //        Console.WriteLine(cellData);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("An error occurred: " + ex.Message);
        //    }
        //}

        


    }
}
