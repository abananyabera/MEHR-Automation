using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class ImportNotChanged
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void Import_Not_Changed(SqlConnection sqlconnection)
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Not_Changed is started \n");
            string Query = "Select distinct(fieldname) from tbl_Employees_Import_Changed_Not_Updated";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                string fieldValue = datareader["fieldname"].ToString();
                Console.WriteLine(fieldValue);

            }
            Console.WriteLine("\n distinct(fieldname) of tbl_Employees_Import_Not_Changed is completed\n");
        }
    }
}
