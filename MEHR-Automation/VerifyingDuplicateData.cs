using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class VerifyingDuplicateData
    {
        ExecuteQueries executeQueries = new ExecuteQueries();


        public void VerifyDuplicateDataintable(SqlConnection sqlconnection)
        {
            #region MyRegion
            Console.WriteLine("\nQuery4 is started\n");
            string Query4 = "select uniqueid,count(uniqueid),datasourceid from tbl_employees_import group by uniqueid,datasourceid having count(uniqueid)>1";
            SqlDataReader datareader4 = executeQueries.ExecuteQuery(Query4, sqlconnection);
            if (!datareader4.HasRows)
            {
                while (datareader4.Read())
                {
                    Console.WriteLine(datareader4[0] + "|" + datareader4[1] + "|" + datareader4[2]);
                }

                Console.WriteLine("\n\nQuery4 is executed\n\n");
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            #endregion


            #region MyRegion
            Console.WriteLine("\nQuery5 is started\n");
            string Query5 = "IF OBJECT_ID('tbl_Employees_Import_Excluded') IS NOT NULL DROP TABLE tbl_Employees_Import_Excluded";
            SqlDataReader datareader5 = executeQueries.ExecuteQuery(Query5, sqlconnection);
            if (!datareader5.HasRows)
            {
                while (datareader5.Read())
                {
                    Console.WriteLine(datareader5[0]);
                }
                Console.WriteLine("\n\nQuery5 is Executed\n\n");
            }
            else
            {

                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            #endregion


            #region MyRegion
            Console.WriteLine("\nQuery6 is started\n");
            string Query6 = "select b.* into tbl_Employees_Import_Excluded from WorkdayIntegratedEmployees a, tbl_employees_import b where me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid";
            SqlDataReader datareader6 = executeQueries.ExecuteQuery(Query6, sqlconnection);
            if (!datareader6.HasRows)
            {
                while (datareader6.Read())
                {
                    Console.WriteLine(datareader6[0] + "|" + datareader6[1]);
                }
                Console.WriteLine("\n\nQuery6 is Executed\n\n");
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            #endregion


            #region MyRegion
            Console.WriteLine("\nQuery7 is started\n");
            string Query7 = "DELETE b from WorkdayIntegratedEmployees a, tbl_employees_import b where me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid";
            SqlDataReader datareader7 = executeQueries.ExecuteQuery(Query7, sqlconnection);
            if (!datareader7.HasRows)
            {

                while (datareader7.Read())
                {
                    Console.WriteLine(datareader7[0] + "|" + datareader7[1]);
                }
                Console.WriteLine("\nQuery7 is Executed\n");
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            #endregion
        }

    }
}
