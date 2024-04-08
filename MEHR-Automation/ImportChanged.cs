using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class ImportChanged
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void Import_Changed(SqlConnection sqlconnection)
        {
            string Query = "Select distinct(fieldname) from tbl_Employees_Import_Changed";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                string fieldValue = datareader["fieldname"].ToString();
                Console.WriteLine(fieldValue);
            }
            Console.WriteLine("distinct(fieldname) of tbl_Employees_Import_Changed is completed");
            
        }
    }
}
