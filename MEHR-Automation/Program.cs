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
                    int select_Import_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_changed)
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

                            

                            //Add data to Excel worksheet
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
                            string excelPath = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            workbook.SaveAs(excelPath);
                            workbook.Close();
                            excelApp.Quit();
                            Console.WriteLine($"Excel file created at: {excelPath}");

                            datareader14.Close();
                            datareader14 = ExecuteQuery(Query14, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath);

                            var peopleReportExcelApp = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook = peopleReportExcelApp.Workbooks.Open(peopleReportPath);
                            var peopleReportWorksheet = (Worksheet)peopleReportWorkbook.Sheets[1];

                            while (datareader14.Read()) //Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader14[0]);
                                var range = peopleReportWorksheet.Range["A:A"];
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, print a message
                                {
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {foundCell.Row}!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook.Close();
                            peopleReportExcelApp.Quit();

                            Console.WriteLine("\n orig_epassid is completed \n");
                            break;

                        case 2:
                            Console.WriteLine("\n orig_internet_email is started \n");
                            string Query15 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_internet_email'";
                            SqlDataReader datareader15 = ExecuteQuery(Query15, sqlconnection);

                            string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp = new Microsoft.Office.Interop.Excel.Application();
                            //existingApp.Visible = true;
                            var existingWorkbook = existingApp.Workbooks.Open(existingPath);

                            // Get or create Sheet2
                            Worksheet sheet2;
                            try
                            {
                                sheet2 = (Worksheet)existingWorkbook.Sheets[2];
                            }
                            catch
                            {
                                // If Sheet2 doesn't exist, add it
                                sheet2 = (Worksheet)existingWorkbook.Sheets.Add(After: existingWorkbook.Sheets[existingWorkbook.Sheets.Count]);
                                sheet2.Name = "orig_internet_email";
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

                            datareader15.Close();
                            datareader15 = ExecuteQuery(Query15, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath2 = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath2);
                            var peopleReportExcelApp2 = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook2 = peopleReportExcelApp2.Workbooks.Open(peopleReportPath2);
                            var peopleReportWorksheet2 = (Worksheet)peopleReportWorkbook2.Sheets[1];
                            while (datareader15.Read()) // Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader15[0]);
                                var range = peopleReportWorksheet2.Range["A:E"]; // Adjust range to cover columns A to E
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, print a message
                                {
                                    var rowinpeoplereport2 = foundCell.Row;
                                    var valueFromColumnE = peopleReportWorksheet2.Cells[rowinpeoplereport2, 5].Value; // Assuming column E is the 3th column (index starts from 1)
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport2}and corresponding value from column C is '{valueFromColumnE}'!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook2.Close();
                            peopleReportExcelApp2.Quit();


                            Console.WriteLine("\n orig_internet_email is completed \n");
                            break;

                        case 3:
                            Console.WriteLine("\n orig_first is started \n");
                            string Query16 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_first'";
                            SqlDataReader datareader16 = ExecuteQuery(Query16, sqlconnection);

                            string existingPath3 = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp3 = new Microsoft.Office.Interop.Excel.Application();
                            //existingApp.Visible = true;
                            var existingWorkbook3 = existingApp3.Workbooks.Open(existingPath3);

                            // Get or create Sheet2
                            Worksheet sheet3;
                            try
                            {
                                sheet3 = (Worksheet)existingWorkbook3.Sheets[3];
                            }
                            catch
                            {
                                // If Sheet3 doesn't exist, add it
                                sheet3 = (Worksheet)existingWorkbook3.Sheets.Add(After: existingWorkbook3.Sheets[existingWorkbook3.Sheets.Count]);
                                sheet3.Name = "orig_first";
                            }

                            // Add column headers
                            for (int i = 0; i < datareader16.FieldCount; i++)
                            {
                                sheet3.Cells[1, i + 1] = datareader16.GetName(i);
                            }


                            // Add data to Sheet2
                            int row3 = 2;
                            while (datareader16.Read())
                            {
                                for (int i = 0; i < datareader16.FieldCount; i++)
                                {
                                    sheet3.Cells[row3, i + 1] = datareader16[i];
                                }
                                row3++;
                            }

                            // Save the existing Excel workbook
                            existingWorkbook3.Save();
                            existingWorkbook3.Close();
                            existingApp3.Quit();

                            Console.WriteLine($"Excel file updated at: {existingPath3}");

                            datareader16.Close();
                            datareader16 = ExecuteQuery(Query16, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath3 = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath3);
                            var peopleReportExcelApp3 = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook3 = peopleReportExcelApp3.Workbooks.Open(peopleReportPath3);
                            var peopleReportWorksheet3 = (Worksheet)peopleReportWorkbook3.Sheets[1];
                            while (datareader16.Read()) //Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader16[0]);
                                var range = peopleReportWorksheet3.Range["A:C"]; // Adjust range to cover columns A to C
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, print a message
                                {
                                    var rowinpeoplereport5 = foundCell.Row;
                                    var valueFromColumnC = peopleReportWorksheet3.Cells[rowinpeoplereport5, 3].Value; // Assuming column C is the 3th column (index starts from 1)
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport5} and corresponding value from column C is '{valueFromColumnC}'!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook3.Close();
                            peopleReportExcelApp3.Quit();

                            Console.WriteLine("\n orig_first is completed \n");

                            Console.WriteLine("\n orig_first is completed \n");
                            break;

                        case 4:
                            Console.WriteLine("\n orig_last is started \n");
                            string Query17 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_last'";
                            SqlDataReader datareader17 = ExecuteQuery(Query17, sqlconnection);

                            string existingPath4 = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp4 = new Microsoft.Office.Interop.Excel.Application();
                            //existingApp.Visible = true;
                            var existingWorkbook4 = existingApp4.Workbooks.Open(existingPath4);

                            // Get or create Sheet2
                            Worksheet sheet4;
                            try
                            {
                                sheet4 = (Worksheet)existingWorkbook4.Sheets[4];
                            }
                            catch
                            {
                                // If Sheet3 doesn't exist, add it
                                sheet4 = (Worksheet)existingWorkbook4.Sheets.Add(After: existingWorkbook4.Sheets[existingWorkbook4.Sheets.Count]);
                                sheet4.Name = "orig_last";
                            }

                            // Add column headers
                            for (int i = 0; i < datareader17.FieldCount; i++)
                            {
                                sheet4.Cells[1, i + 1] = datareader17.GetName(i);
                            }


                            // Add data to Sheet4
                            int row4 = 2;
                            while (datareader17.Read())
                            {
                                for (int i = 0; i < datareader17.FieldCount; i++)
                                {
                                    sheet4.Cells[row4, i + 1] = datareader17[i];
                                }
                                row4++;
                            }

                            // Save the existing Excel workbook
                            existingWorkbook4.Save();
                            existingWorkbook4.Close();
                            existingApp4.Quit();

                            Console.WriteLine($"Excel file updated at: {existingPath4}");

                            datareader17.Close();
                            datareader17 = ExecuteQuery(Query17, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath4 = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath4);
                            var peopleReportExcelApp4 = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook4 = peopleReportExcelApp4.Workbooks.Open(peopleReportPath4);
                            var peopleReportWorksheet4 = (Worksheet)peopleReportWorkbook4.Sheets[1];
                            while (datareader17.Read()) // Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader17[0]); // Assuming the index is 0, change it if needed
                                var range = peopleReportWorksheet4.Range["A:D"]; // Adjust range to cover columns A to D
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, retrieve corresponding value from column D
                                {
                                    var rowinpeoplereport4 = foundCell.Row;
                                    var valueFromColumnD = peopleReportWorksheet4.Cells[rowinpeoplereport4, 4].Value; // Assuming column D is the 4th column (index starts from 1)
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport4} and corresponding value from column D is '{valueFromColumnD}'!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook4.Close();
                            peopleReportExcelApp4.Quit();

                            Console.WriteLine("\n orig_last is completed \n");
                            break;

                        case 5:
                            Console.WriteLine("\n orig_middle is started \n");
                            string Query18 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_middle'";
                            SqlDataReader datareader18 = ExecuteQuery(Query18, sqlconnection);

                            string existingPath5 = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp5 = new Microsoft.Office.Interop.Excel.Application();
                            //existingApp.Visible = true;
                            var existingWorkbook5 = existingApp5.Workbooks.Open(existingPath5);

                            // Get or create Sheet2
                            Worksheet sheet5;
                            try
                            {
                                sheet5 = (Worksheet)existingWorkbook5.Sheets[5];
                            }
                            catch
                            {
                                // If Sheet3 doesn't exist, add it
                                sheet5 = (Worksheet)existingWorkbook5.Sheets.Add(After: existingWorkbook5.Sheets[existingWorkbook5.Sheets.Count]);
                                sheet5.Name = "orig_middle";
                            }

                            // Add column headers
                            for (int i = 0; i < datareader18.FieldCount; i++)
                            {
                                sheet5.Cells[1, i + 1] = datareader18.GetName(i);
                            }


                            // Add data to Sheet4
                            int row5 = 2;
                            while (datareader18.Read())
                            {
                                for (int i = 0; i < datareader18.FieldCount; i++)
                                {
                                    sheet5.Cells[row5, i + 1] = datareader18[i];
                                }
                                row5++;
                            }

                            // Save the existing Excel workbook
                            existingWorkbook5.Save();
                            existingWorkbook5.Close();
                            existingApp5.Quit();

                            Console.WriteLine($"Excel file updated at: {existingPath5}");

                            datareader18.Close();
                            datareader18 = ExecuteQuery(Query18, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath5 = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath5);
                            var peopleReportExcelApp5 = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook5 = peopleReportExcelApp5.Workbooks.Open(peopleReportPath5);
                            var peopleReportWorksheet5 = (Worksheet)peopleReportWorkbook5.Sheets[1];
                            while (datareader18.Read()) // Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader18[0]);
                                var range = peopleReportWorksheet5.Range["A:G"]; // Adjust range to cover columns A to G
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, print a message
                                {
                                    var rowinpeoplereport5 = foundCell.Row;
                                    var valueFromColumnE = peopleReportWorksheet5.Cells[rowinpeoplereport5, 7].Value; // Assuming column E is the 7th column (index starts from 1)
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport5}and corresponding value from column E is '{valueFromColumnE}'!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook5.Close();
                            peopleReportExcelApp5.Quit();

                            Console.WriteLine("\n orig_middle is completed \n");
                            break;

                        default:
                            Environment.Exit(0);
                            break;

                    }
                    count++;
                }


                Console.ReadLine();

                Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Not_Changed is started \n");
                string Query19 = "Select distinct(fieldname) from tbl_Employees_Import_Changed_Not_Updated";
                SqlDataReader datareader19 = ExecuteQuery(Query19, sqlconnection);
                while (datareader19.Read())
                {
                    string fieldValue = datareader19["fieldname"].ToString();
                    Console.WriteLine(fieldValue);

                }
                Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Not_Changed is completed\n");



                int count1 = 0;
                while (count1 < 1)
                {
                    Console.WriteLine("select the fieldname that need to be executed  \n 1.orig_internet_email ");
                    int select_Import_changed = int.Parse(Console.ReadLine());
                    switch (select_Import_changed)
                    {
                        case 1:
                            Console.WriteLine("\n orig_internet_email_Import_Not_Changed is started \n");
                            string Query15 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed_Not_Updated A inner join dbo.tbl_employees_stage1 B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_internet_email'";
                            SqlDataReader datareader15 = ExecuteQuery(Query15, sqlconnection);

                            string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel.xlsx";
                            Microsoft.Office.Interop.Excel.Application existingApp = new Microsoft.Office.Interop.Excel.Application();
                            //existingApp.Visible = true;
                            var existingWorkbook = existingApp.Workbooks.Open(existingPath);

                            // Get or create Sheet2
                            Worksheet sheet2;
                            try
                            {
                                sheet2 = (Worksheet)existingWorkbook.Sheets[6];
                            }
                            catch
                            {
                                // If Sheet2 doesn't exist, add it
                                sheet2 = (Worksheet)existingWorkbook.Sheets.Add(After: existingWorkbook.Sheets[existingWorkbook.Sheets.Count]);
                                sheet2.Name = "itert_email_ImptNotChangd";
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

                            datareader15.Close();
                            datareader15 = ExecuteQuery(Query15, sqlconnection);

                            // Search in People_Report_1215.xlsx
                            string peopleReportPath2 = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                            Console.WriteLine(peopleReportPath2);
                            var peopleReportExcelApp2 = new Microsoft.Office.Interop.Excel.Application();
                            var peopleReportWorkbook2 = peopleReportExcelApp2.Workbooks.Open(peopleReportPath2);
                            var peopleReportWorksheet2 = (Worksheet)peopleReportWorkbook2.Sheets[1];
                            while (datareader15.Read()) // Iterate over each value in datareader[0] and perform the search
                            {
                                var searchValue = Convert.ToString(datareader15[0]);
                                var range = peopleReportWorksheet2.Range["A:E"]; // Adjust range to cover columns A to E
                                var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                                if (foundCell != null) // If the value is found, print a message
                                {
                                    var rowinpeoplereport2 = foundCell.Row;
                                    var valueFromColumnE = peopleReportWorksheet2.Cells[rowinpeoplereport2, 5].Value; // Assuming column E is the 3th column (index starts from 1)
                                    Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport2}and corresponding value from column C is '{valueFromColumnE}'!");
                                }
                                else
                                {
                                    Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                                }
                            }

                            // Close the workbook and quit Excel application
                            peopleReportWorkbook2.Close();
                            peopleReportExcelApp2.Quit();
                            Console.WriteLine("Org_internet_Email_Import_Not_Changed is complelted");
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
                storedProcedure.procUpdatetempfixStoredProcedure(sqlconnection);
                storedProcedure.procPostUpdateStoredProcedure(sqlconnection);  //  have to check this..

                //MACROQUERIES EXECUTION
                MacroQueries macroQueries = new MacroQueries();
                macroQueries.Add_Counts_by_Datasource(sqlconnection); Console.ReadLine();
                macroQueries.Changed_Fields_by_DataSource(sqlconnection); Console.ReadLine();
                macroQueries.Change_NotUpdated_by_DataSource(sqlconnection); Console.ReadLine();
                macroQueries.AddDeleted_by_DataSource(sqlconnection); Console.ReadLine();
                macroQueries.Removed_Countby_DataSource(sqlconnection); Console.ReadLine();
                macroQueries.Check_Email_Types(sqlconnection); Console.ReadLine();
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


//        Open the people's report Excel file
//        string peopleReportPath = "C:\\Path\\To\\PeopleReport.xlsx";
//        var peopleReportExcelApp = new Microsoft.Office.Interop.Excel.Application();
//        var peopleReportWorkbook = peopleReportExcelApp.Workbooks.Open(peopleReportPath);
//        var peopleReportWorksheet = (Worksheet)peopleReportWorkbook.Sheets[1];

//        Get the range of values in column A of the people's report Excel sheet
//        var range = peopleReportWorksheet.Range["A:A"];

//        Check if datareader14[0] is present in column A
//if (range.Cells.Find(Convert.ToString(datareader14[0]), Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole) != null)
//{
//    Console.WriteLine("Hello, the value is present in the people's report!");
//}

//    Close the people's report Excel file
//    peopleReportWorkbook.Close();
//peopleReportExcelApp.Quit();
        


    }
}
