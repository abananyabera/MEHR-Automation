using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class OrigEpassId
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void execQuery(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n orig_epassid is started \n");
            string Query14 = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from dbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid where a.fieldname = 'orig_epassid'";
            SqlDataReader datareader14 = executeQueries.ExecuteQuery(Query14, sqlconnection);

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
            string Pathname = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
            workbook.SaveAs(Pathname);
            workbook.Close();
            excelApp.Quit();
            //Console.WriteLine($"Excel file created at: {excelPath}");

            datareader14.Close();
            datareader14 = executeQueries.ExecuteQuery(Query14, sqlconnection);
                
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
           
        }
    }
}
