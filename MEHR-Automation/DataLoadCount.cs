using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class DataLoadCount
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        public bool Dataloadfile(SqlConnection sqlconnection)
        {
            //Read csv file Dataload_77 csv file
            int lineCount = CountLinesInCsvFile(@"C:\Users\kwr579\Desktop\AUTOMATION\DataLoad_77_AfterTriage.csv");
            Console.WriteLine("Number of lines in the file: " + lineCount);
            int comparisioncount = 0;
            string Query3 = "Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query3, sqlconnection);
            while (datareader.Read())
            {
                comparisioncount = comparisioncount + (int)datareader[0];
                Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
            }
            Console.WriteLine(comparisioncount);
            if (lineCount == comparisioncount)
            {
                return true;
            }
            else
            {             
                return false;
            }
        }
        public  int CountLinesInCsvFile(string filepath)
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
