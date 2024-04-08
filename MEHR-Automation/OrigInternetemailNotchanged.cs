using Microsoft.Office.Interop.Excel;
using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigInternetemailNotchanged
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void Orig_internet_Email_Not_changed(SqlConnection sqlconnection)
        {
            Console.WriteLine("orig_internet_email_Import_Not_Changed is started");
            string Query = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed_Not_Updated A inner join dbo.tbl_employees_stage1 B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_internet_email'";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (datareader.HasRows)
            {

                string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
                Microsoft.Office.Interop.Excel.Application existingApp = new Microsoft.Office.Interop.Excel.Application();
                //existingApp.Visible = true;
                var existingWorkbook = existingApp.Workbooks.Open(existingPath);

                // Get or create Sheet2
                Worksheet sheet;
                try
                {
                    sheet = (Worksheet)existingWorkbook.Sheets[6];
                }
                catch
                {
                    // If Sheet2 doesn't exist, add it
                    sheet = (Worksheet)existingWorkbook.Sheets.Add(After: existingWorkbook.Sheets[existingWorkbook.Sheets.Count]);
                    sheet.Name = "itert_email_ImptNotChangd";
                }

                // Add column headers
                for (int i = 0; i < datareader.FieldCount; i++)
                {
                    sheet.Cells[1, i + 1] = datareader.GetName(i);
                }


                // Add data to Sheet2
                int row2 = 2;
                while (datareader.Read())
                {
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        sheet.Cells[row2, i + 1] = datareader[i];
                    }
                    row2++;
                }

                // Save the existing Excel workbook
                existingWorkbook.Save();
                existingWorkbook.Close();
                existingApp.Quit();

                Console.WriteLine($"Excel file updated at: {existingPath}");

                datareader.Close();
                datareader = executeQueries.ExecuteQuery(Query, sqlconnection);

                // Search in People_Report_1215.xlsx
                string peopleReportPath = @userProfileDirectory + "\\AUTOMATION\\People_Report_1215.xlsx";
                var peopleReportExcelApp = new Microsoft.Office.Interop.Excel.Application();
                var peopleReportWorkbook = peopleReportExcelApp.Workbooks.Open(peopleReportPath);
                var peopleReportWorksheet = (Worksheet)peopleReportWorkbook.Sheets[1];
                while (datareader.Read()) // Iterate over each value in datareader[0] and perform the search
                {
                    var searchValue = Convert.ToString(datareader[0]);
                    var internet_email = Convert.ToString(datareader[9]);
                    var range = peopleReportWorksheet.Range["A:E"]; // Adjust range to cover columns A to E
                    var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                    if (foundCell != null) // If the value is found, print a message
                    {
                        var rowinpeoplereport = foundCell.Row;
                        var valueFromColumnE = peopleReportWorksheet.Cells[rowinpeoplereport, 5].Value; // Assuming column E is the 3th column (index starts from 1)
                        Console.WriteLine($"\nThe value '{searchValue}' is present in the people's report at row {rowinpeoplereport}and corresponding value from column C is '{valueFromColumnE}'!");
                        if (internet_email == valueFromColumnE)
                        {
                            Console.WriteLine("No Update is Required");
                        }
                        else
                        {
                            Console.WriteLine("Update Required on the Org_internet_email for Import not changed");
                            string Orig_internet_email_Update = "Update tbl_employees_stage1 set internet_email = orig_internet_email where epassid in ('" + datareader[0] + "')";
                            SqlDataReader datareader_Update_internetemail = executeQueries.ExecuteQuery(Orig_internet_email_Update, sqlconnection);
                            Console.WriteLine("Org_internet_email for the import not changed is updated");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"\nThe value '{searchValue}' is not present in the people's report.");
                    }
                }

                // Close the workbook and quit Excel application
                peopleReportWorkbook.Close();
                peopleReportExcelApp.Quit();
                Console.WriteLine("Org_internet_Email_Import_Not_Changed is complelted");
                ReadLine();
            }
            else
            {
                Console.WriteLine("\n orig_internet_email is completed with out generating the Excel because of Query Returning Empty.");
                ReadLine();
            }
        }
        static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");

        }
    }
}
