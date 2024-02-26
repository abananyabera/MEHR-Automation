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
            
            Console.WriteLine("\n count of tbl_Employees_Import_Add is started ");
            string Import_Add_count_Query = "select count (*) from tbl_Employees_Import_Add";
            SqlDataReader Import_Add_datareader = executeQueries.ExecuteQuery(Import_Add_count_Query, sqlconnection);
            while (Import_Add_datareader.Read())
            {
                int count = Convert.ToInt32(Import_Add_datareader[0]);
                Console.WriteLine("count of tbl_Employees_Import_Add : " + Import_Add_datareader[0]);
                if (count > 500)
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Add is greater than 500 we can't procedd further please reach out to the workday team for the confirmartion");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Add is completed");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
            }
            
            
            Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is started ");
            string Import_Add_Deleted_Query = "select count (*) from tbl_Employees_Import_Add_Deleted";
            SqlDataReader Import_Add_Deleted_datareader = executeQueries.ExecuteQuery(Import_Add_Deleted_Query, sqlconnection);
            while (Import_Add_Deleted_datareader.Read())
            {
                int count = Convert.ToInt32(Import_Add_Deleted_datareader[0]);
                Console.WriteLine("Count of tbl_Employees_Import_Add_Deleted : " + Import_Add_Deleted_datareader[0]);
                if (count > 500)
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is greater than 500 we can't procedd further please reach out to the workday team for the confirmartion");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Add_Deleted is completed");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
            }
            
            

            #region MyRegion
            Console.WriteLine("\n count of tbl_Employees_Import_Remove is started ");
            string Import_Remove_Query = "select count (*) from tbl_Employees_Import_Remove";
            SqlDataReader Import_Remove_Query_datareader = executeQueries.ExecuteQuery(Import_Remove_Query, sqlconnection);
            while (Import_Remove_Query_datareader.Read())
            {
                int count = Convert.ToInt32(Import_Remove_Query_datareader[0]);
                Console.WriteLine("Count of tbl_Employees_Import_Remove : " + Import_Remove_Query_datareader[0]);
                if (count > 500)
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Remove is greater than 500 we can't procedd further please reach out to the workday team for the confirmartion");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
                else
                {
                    Console.WriteLine("\n count of tbl_Employees_Import_Remove is completed ");
                    Console.WriteLine("-------------------------------------------------------------");
                    Console.ReadLine();
                }
            }
            
            #endregion
        }

    }
}
