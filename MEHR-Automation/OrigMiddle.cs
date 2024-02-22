using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigMiddle
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void execQuery(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n orig_middle is started \n");
            string Query18 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_middle'";
            SqlDataReader datareader18 = executeQueries.ExecuteQuery(Query18, sqlconnection);

            string existingPath5 = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
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
            datareader18 = executeQueries.ExecuteQuery(Query18, sqlconnection);

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
        }
    }
}
