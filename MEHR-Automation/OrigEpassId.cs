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
        public void OrigEpassId_Query(SqlConnection sqlconnection)
        {
            Console.WriteLine("orig_epassid is started");
            string Query = "Select a.orig_epassid,b.epassid,a.orig_first,b.first,a.orig_last,b.last,a.orig_middle,b.middle,a.orig_internet_email,b.internet_email,a.orig_site,b.site from dbo.tbl_Employees_Import_Changed A inner join dbo.tbl_employees_stage1_Hold B on A.masterid = b.masterid where a.fieldname = 'orig_epassid'";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (datareader.HasRows)
            {

                //create excel workbook
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excelApp.Workbooks.Add();
                var worksheet = (Worksheet)workbook.Sheets[1];

                //Add column headers
                for (int i = 0; i < datareader.FieldCount; i++)
                {
                    worksheet.Cells[1, i + 1] = datareader.GetName(i);
                }



                //Add data to Excel worksheet
                int row = 2;
                while (datareader.Read())
                {
                    for (int i = 0; i < datareader.FieldCount; i++)
                    {
                        worksheet.Cells[row, i + 1] = datareader[i];
                    }
                    row++;
                }

                // Save Excel workbook
                string Pathname = @userProfileDirectory + "\\AUTOMATION\\Excel1.xlsx";
                workbook.SaveAs(Pathname);
                workbook.Close();
                excelApp.Quit();

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
                    var orig_epassid = Convert.ToString(datareader[0]);
                    var range = peopleReportWorksheet.Range["A:A"];
                    var foundCell = range.Cells.Find(searchValue, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                    if (foundCell != null) // If the value is found, print a message
                    {
                        var rowinpeoplereport = foundCell.Row;
                        var valueFromColumnA = peopleReportWorksheet.Cells[rowinpeoplereport, 1].Value; // Assuming column A is the 1st column (index starts from 1)
                        Console.WriteLine($"\nThe value '{searchValue}' is present in the people's report at row {rowinpeoplereport} and corresponding value from column A is '{valueFromColumnA}'!");
                        if (orig_epassid == valueFromColumnA)
                        {
                            Console.WriteLine("No Update is Required");
                        }
                        else
                        {
                            Console.WriteLine("Update Required on the Org_epassid");
                            string Orig_epassid_Update = "Update Stage1 set Stage1.epassid = hold.epassid\r\nfrom tbl_employees_stage1 as stage1\r\njoin tbl_Employees_Stage1_Hold hold on stage1.masterid = hold.masterid \r\nwhere stage1.epassid in ('" + datareader[0] + "')";
                            SqlDataReader datareader_Update_epassid = executeQueries.ExecuteQuery(Orig_epassid_Update, sqlconnection);
                            Console.WriteLine("Org_epassid is updated");

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

                Console.WriteLine("\n orig_epassid is completed ");
                ReadLine();
            }
            else
            {
                Console.WriteLine("\n orig_epassid is completed with out generating the Excel because of Query Returning Empty.");
                ReadLine();
            }
           
        }
        static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");

        }
    }

    
}
