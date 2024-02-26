using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigFirst
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void execQuery(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n orig_first is started \n");
            string Query = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_first'";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (datareader.HasRows)
            {

                string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
                Microsoft.Office.Interop.Excel.Application existingApp = new Microsoft.Office.Interop.Excel.Application();
                //existingApp.Visible = true;
                var existingWorkbook = existingApp.Workbooks.Open(existingPath);

                // Get or create Sheet3
                Worksheet sheet;
                try
                {
                    sheet = (Worksheet)existingWorkbook.Sheets[3];
                }
                catch
                {
                    // If Sheet3 doesn't exist, add it
                    sheet = (Worksheet)existingWorkbook.Sheets.Add(After: existingWorkbook.Sheets[existingWorkbook.Sheets.Count]);
                    sheet.Name = "orig_first";
                }

                // Add column headers
                for (int i = 0; i < datareader.FieldCount; i++)
                {
                    sheet.Cells[1, i + 1] = datareader.GetName(i);
                }


                // Add data to Sheet3
                int row = 2;
                while (datareader.Read())
                {
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        sheet.Cells[row, i + 1] = datareader[i];
                    }
                    row++;
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
                while (datareader.Read()) //Iterate over each value in datareader[0] and perform the search
                {
                    var searchValue = Convert.ToString(datareader[0]);
                    var orig_First = Convert.ToString(datareader[2]);
                    var range = peopleReportWorksheet.Range["A:C"]; // Adjust range to cover columns A to C
                    var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                    if (foundCell != null) // If the value is found, print a message
                    {
                        var rowinpeoplereport = foundCell.Row;
                        var valueFromColumnC = peopleReportWorksheet.Cells[rowinpeoplereport, 3].Value; // Assuming column C is the 3th column (index starts from 1)
                        Console.WriteLine($"The value '{searchValue}' is present in the people's report at row {rowinpeoplereport} and corresponding value from column C is '{valueFromColumnC}'!");
                        if (orig_First == valueFromColumnC)
                        {
                            Console.WriteLine("No Update is Required");
                        }
                        else
                        {
                            Console.WriteLine("Update Required on the Org_First");
                            string Orig_First_Update = "Update Stage1 set Stage1.first = hold.first\r\nfrom tbl_employees_stage1 as stage1\r\njoin tbl_Employees_Stage1_Hold hold on stage1.masterid = hold.masterid \r\nwhere stage1.epassid in ('" + datareader[0] + "')";
                            SqlDataReader datareader_Update_First = executeQueries.ExecuteQuery(Orig_First_Update, sqlconnection);
                            Console.WriteLine("Org_First is updated");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"The value '{searchValue}' is not present in the people's report.");
                    }
                }

                // Close the workbook and quit Excel application
                peopleReportWorkbook.Close();
                peopleReportExcelApp.Quit();

                Console.WriteLine("\n orig_first is completed \n");
            }
            else
            {
                Console.WriteLine("\n orig_first is completed with out generating the Excel because of Query Returning Empty.");
            }

        }
    }
}
