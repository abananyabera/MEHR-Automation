using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigLast
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void execQuery(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n orig_last is started \n");
            string Query17 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_Employees_Stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_last'";
            SqlDataReader datareader17 = executeQueries.ExecuteQuery(Query17, sqlconnection);

            string existingPath4 = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
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
            datareader17 = executeQueries.ExecuteQuery(Query17, sqlconnection);

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
        }
    }
}
