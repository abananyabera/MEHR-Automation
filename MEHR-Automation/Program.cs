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
            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            sqlconnection.Open();
            Console.WriteLine("Connection is successfull");
            //execute commands
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            #region MyRegion

            tablebackup tableBackup = new tablebackup();
            tableBackup.takeTableBackup(sqlconnection);

            //string timeStamp = DateTime.Now.ToString("MMddyyyy");
            //string destinationTable1 = "[dbo]. [tbl_employees_stage1_" + timeStamp + "]";
            //string query1 = "select * into" + " " + destinationTable1 + " " + "from [dbo]. [tbl_employees_stage1]";


            //// checking the table is already presnet or not if present returns 1 else return 0
            //int connectionresult = 0;
            //string checkingtable = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_" + timeStamp + "'";
            //SqlDataReader connection = ExecuteQuery(checkingtable, sqlconnection);
            //while (connection.Read())
            //{
            //    connectionresult = (int)connection[0];
            //}

            ////drop backup table if already present
            //if (connectionresult == 1)
            //{
            //    SqlCommand cmd = new SqlCommand("drop table " + destinationTable1 , sqlconnection);
            //    cmd.ExecuteNonQuery();
            //    Console.WriteLine(destinationTable1 + "dropped succesfully");
            //}

            //// creates the backuptable 1 if not present else it returns error
            //ExecuteQuery(query1, sqlconnection);


            //int countMainTable = 0;
            //string query4 = "Select count(*) from [dbo]. [tbl_employees_stage1]";
            //SqlDataReader counter0 = ExecuteQuery(query4, sqlconnection);
            //while (counter0.Read())
            //{
            //    countMainTable = (int)counter0[0];
            //}

            //int countMainTableBackup = 0;
            //string query5 = "Select count(*) from " + destinationTable1;
            //SqlDataReader counter1 = ExecuteQuery(query5, sqlconnection);
            //while (counter1.Read())
            //{
            //    countMainTableBackup = (int)counter1[0];
            //}

            //if (countMainTable == countMainTableBackup)
            //{
            //    Console.WriteLine("Backup for tbl_employees_stage1 is successful");
            //}
            //else
            //{
            //    Console.WriteLine("Backup for tbl_employees_stage1 is failed");
            //}
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            #endregion

            #region MyRegion

            //string timeStamp2 = DateTime.Now.ToString("MMddyyyy");
            //string destinationTable2 = "[dbo]. [tbl_employees_stage1_hold_" + timeStamp2 + "]";
            //string backupquery = "select * into" + " " + destinationTable2 + " " + "from [dbo]. [tbl_employees_stage1_hold]";


            //// checking the table is already presnet or not if present returns 1 else return 0
            //int connectionresult2 = 0;
            //string checkingtable2 = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_hold_" + timeStamp2 + "'";
            
            //SqlDataReader connection2 = ExecuteQuery(checkingtable2, sqlconnection);
            //while (connection2.Read())
            //{
            //    connectionresult2 = (int)connection2[0];
                
            //}

            ////drop backup table 2 if already present
            //if (connectionresult2 == 1)
            //{
            //    SqlCommand cmd = new SqlCommand("drop table " + destinationTable2, sqlconnection);
                
            //    cmd.ExecuteNonQuery();
            //    Console.WriteLine(destinationTable2 + "dropped succesfully");
            //}

            //// creates the backuptable 2 if not present else it returns error
            //ExecuteQuery(backupquery, sqlconnection);


            //int countMainTable2 = 0;
            //string backupquery4 = "Select count(*) from tbl_employees_stage1_hold";
            //SqlDataReader backupcounter1 = ExecuteQuery(backupquery4, sqlconnection);
            //while (counter0.Read())
            //{
            //    countMainTable2 = (int)backupcounter1[0];
            //}

            //int countMainTableBackup2 = 0;
            //string backquery5 = "Select count(*) from " + destinationTable2;
            //SqlDataReader backupcounter2 = ExecuteQuery(backquery5, sqlconnection);
            //while (counter1.Read())
            //{
            //    countMainTableBackup2 = (int)backupcounter2[0];
            //}

            //if (countMainTable2 == countMainTableBackup2)
            //{
            //    Console.WriteLine("Backup for tbl_employees_stage1_hold is successful");
            //}
            //else
            //{
            //    Console.WriteLine("Backup for tbl_employees_stage1_hold is failed");
            //}
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            #endregion

            //Read csv file Dataload_77 csv file
            int lineCount = CountLinesInCsvFile(@"C:\Users\kwr579\Desktop\AUTOMATION\DataLoad_77_AfterTriage.csv");
            Console.WriteLine("Number of lines in the file: " + lineCount);

            #region MyRegion
            int comparisioncount = 0;
            string Query3 = "Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid";
            SqlDataReader datareader = ExecuteQuery(Query3, sqlconnection);
            while (datareader.Read())
            {
                
                comparisioncount = comparisioncount +(int)datareader[0];
                Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
                

            }
            Console.WriteLine(comparisioncount);
            if (lineCount == comparisioncount)
            {
                Console.WriteLine("Tharun Sai");
            }
            else
            {
                Console.WriteLine("Kanna");
            }
            Console.WriteLine("------------------Query3 is executed----------------");
            #endregion

            Console.WriteLine("Please Disable the jobs: \n 1. UserID Synch Job \n 2. Modification Synch Job \n Please press Y to continue");
            char chkJobs = Console.ReadLine()[0];
            

            if (chkJobs == 'Y' || chkJobs == 'y')
            {
                Console.WriteLine("\n\n\nJob disabled successfully\n\n\n");

                #region MyRegion
                Console.WriteLine("\nQuery4 is started\n");
                string Query4 = "select uniqueid,count(uniqueid),datasourceid from tbl_employees_import group by uniqueid,datasourceid having count(uniqueid)>1";
                SqlDataReader datareader4 = ExecuteQuery(Query4, sqlconnection);
                if (!datareader4.HasRows)
                {
                    while (datareader4.Read())
                    {
                        Console.WriteLine(datareader4[0] + "|" + datareader4[1] + "|" + datareader4[2]);
                    }

                    Console.WriteLine("\n\nQuery4 is executed\n\n");
                }
                else
                {
                    Console.WriteLine("Error");
                    Console.ReadLine();
                }
                #endregion


                #region MyRegion
                Console.WriteLine("\nQuery5 is started\n");
                string Query5 = "IF OBJECT_ID('tbl_Employees_Import_Excluded') IS NOT NULL DROP TABLE tbl_Employees_Import_Excluded";
                SqlDataReader datareader5= ExecuteQuery(Query5, sqlconnection);
                if (!datareader5.HasRows)
                {
                    while (datareader5.Read())
                    {
                        Console.WriteLine(datareader5[0]);
                    }
                    Console.WriteLine("\n\nQuery5 is Executed\n\n");
                }
                else 
                { 

                    Console.WriteLine("Error");
                    Console.ReadLine();
                }
                #endregion


                #region MyRegion
                Console.WriteLine("\nQuery6 is started\n");
                string Query6 = "select b.* into tbl_Employees_Import_Excluded from WorkdayIntegratedEmployees a, tbl_employees_import b where me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid";
                SqlDataReader datareader6 = ExecuteQuery(Query6, sqlconnection);
                if (!datareader6.HasRows)
                {
                    while (datareader6.Read())
                    {
                        Console.WriteLine(datareader6[0] + "|" + datareader6[1]);
                    }
                    Console.WriteLine("\n\nQuery6 is Executed\n\n");
                }
                else
                {
                    Console.WriteLine("Error");
                    Console.ReadLine();
                }
                #endregion


                #region MyRegion
                Console.WriteLine("\nQuery7 is started\n");
                string Query7 = "DELETE b from WorkdayIntegratedEmployees a, tbl_employees_import b where me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid";
                SqlDataReader datareader7 = ExecuteQuery(Query7, sqlconnection);
                if (!datareader7.HasRows)
                {

                    while (datareader7.Read())
                    {
                        Console.WriteLine(datareader7[0] + "|" + datareader7[1]);
                    }
                    Console.WriteLine("\nQuery7 is Executed\n");
                }
                else
                {
                    Console.WriteLine("Error");
                    Console.ReadLine();
                }
                #endregion

                #region MyRegion
                //Console.WriteLine("\n stored procedure proc_Pre_Update_Processing started\n");
                //string Query8 = "exec proc_Pre_Update_Processing";
                //SqlDataReader datareader8 = ExecuteQuery(Query8, sqlconnection);
                //while (datareader8.Read())
                //{
                //    Console.WriteLine(datareader8[0] + "|" + datareader8[1]);
                //}
                //Console.WriteLine("\nstored procedure proc_Pre_Update_Processing is Executed\n");
                #endregion

                #region MyRegion
                //Console.WriteLine("\n stored procedure procProcessEmployeeUpdates started\n");
                //string Query9 = "exec procProcessEmployeeUpdates";
                //SqlDataReader datareader9 = ExecuteQuery(Query9, sqlconnection);
                //while (datareader9.Read())
                //{
                //    Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
                //}
                //Console.WriteLine("\nstored procedure procProcessEmployeeUpdates is Executed\n");
                #endregion

                #region MyRegion
                Console.WriteLine("\n count of tbl_Employees_Import_Add is started \n");
                string Query10 = "select count (*) from tbl_Employees_Import_Add";
                SqlDataReader datareader10 = ExecuteQuery(Query10, sqlconnection);
                while (datareader10.Read())
                {   
                    Console.WriteLine(datareader10[0]);
                }
                Console.WriteLine("\n count of tbl_Employees_Import_Add is completed \n");
               
                #endregion


                #region MyRegion
                Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is started \n");
                string Query11 = "select count (*) from tbl_Employees_Import_Add_Deleted";
                SqlDataReader datareader11 = ExecuteQuery(Query11, sqlconnection);
                while (datareader11.Read())
                {
                    Console.WriteLine(datareader11[0]);
                }
                Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is completed \n");
                #endregion

                #region MyRegion
                Console.WriteLine("\n count of tbl_Employees_Import_Remove is started \n");
                string Query12 = "select count (*) from tbl_Employees_Import_Remove";
                SqlDataReader datareader12 = ExecuteQuery(Query12, sqlconnection);
                while (datareader12.Read())
                {
                    Console.WriteLine(datareader12[0]);
                }
                Console.WriteLine("\n count of tbl_Employees_Import_Remove is completed \n");
                #endregion

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
                            string Pathname = "C:\\Users\\kwr579\\Desktop\\AUTOMATION\\Excel1.xlsx";
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

                            string existingPath = "C:\\Users\\kwr579\\Desktop\\AUTOMATION\\Excel1.xlsx";
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
                Console.WriteLine("You have not disabled the job.");
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

        public static int CountLinesInCsvFile(string filepath)
        {
            int count = 0;

            try
            {
                using (StreamReader reader = new StreamReader(filepath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        count++;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return count;
        }


    }
}
