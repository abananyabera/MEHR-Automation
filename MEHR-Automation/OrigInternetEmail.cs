using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigInternetEmail
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void execQuery (SqlConnection sqlconnection)
        {
            Console.WriteLine("\n orig_internet_email is started \n");
            string Query15 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from \r\ndbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid\r\nwhere a.fieldname = 'orig_internet_email'";
            SqlDataReader datareader15 = executeQueries.ExecuteQuery(Query15, sqlconnection);

            string existingPath = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
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
            datareader15 = executeQueries.ExecuteQuery(Query15, sqlconnection);

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
        }
    }
}
