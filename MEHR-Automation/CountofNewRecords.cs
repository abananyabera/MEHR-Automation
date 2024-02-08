using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class CountofNewRecords
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void CountNewRecords(SqlConnection sqlconnection)
        {
            #region MyRegion
            Console.WriteLine("\n count of tbl_Employees_Import_Add is started \n");
            string Query10 = "select count (*) from tbl_Employees_Import_Add";
            SqlDataReader datareader10 = executeQueries.ExecuteQuery(Query10, sqlconnection);
            while (datareader10.Read())
            {
                Console.WriteLine(datareader10[0]);
            }
            Console.WriteLine("\n count of tbl_Employees_Import_Add is completed \n");

            #endregion


            #region MyRegion
            Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is started \n");
            string Query11 = "select count (*) from tbl_Employees_Import_Add_Deleted";
            SqlDataReader datareader11 = executeQueries.ExecuteQuery(Query11, sqlconnection);
            while (datareader11.Read())
            {
                Console.WriteLine(datareader11[0]);
            }
            Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is completed \n");
            #endregion

            #region MyRegion
            Console.WriteLine("\n count of tbl_Employees_Import_Remove is started \n");
            string Query12 = "select count (*) from tbl_Employees_Import_Remove";
            SqlDataReader datareader12 = executeQueries.ExecuteQuery(Query12, sqlconnection);
            while (datareader12.Read())
            {
                Console.WriteLine(datareader12[0]);
            }
            Console.WriteLine("\n count of tbl_Employees_Import_Remove is completed \n");
            #endregion
        }

    }
}
