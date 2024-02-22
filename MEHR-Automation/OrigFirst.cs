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
            string Query16 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_first'";
            SqlDataReader datareader16 = executeQueries.ExecuteQuery(Query16, sqlconnection);

            string existingPath3 = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
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
            datareader16 = executeQueries.ExecuteQuery(Query16, sqlconnection);

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

        }
    }
}
