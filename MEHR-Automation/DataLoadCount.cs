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

            string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            //Read csv file Dataload_77 csv file2
            int lineCount = CountLinesInCsvFile(@userProfileDirectory+ "\\AUTOMATION\\DataLoad_77_AfterTriage.csv");
            Console.WriteLine("path of the triage file  : " + userProfileDirectory + "\\AUTOMATION\\DataLoad_77_AfterTriage.csv");
             
            Console.WriteLine("Number of lines in the Triage file: " + lineCount);
            int comparisioncount = 0;
            string Query = "Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                comparisioncount = comparisioncount + (int)datareader[0];
                
            }
            Console.WriteLine("Count of the Query Result :" +comparisioncount);
            
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
