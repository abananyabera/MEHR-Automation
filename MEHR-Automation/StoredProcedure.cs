using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class StoredProcedure
    {
        ExecuteQueries executeQueries = new ExecuteQueries();

        public void StoredProcedureExecution(SqlConnection sqlconnection)
        {
            #region MyRegion
            //Console.WriteLine("\n stored procedure proc_Pre_Update_Processing started\n");
            //string Query8 = "exec proc_Pre_Update_Processing";
            //SqlDataReader datareader8 = executeQueries.ExecuteQuery(Query8, sqlconnection);
            //while (datareader8.Read())
            //{
            //    Console.WriteLine(datareader8[0] + "|" + datareader8[1]);
            //}
            //Console.WriteLine("\nstored procedure proc_Pre_Update_Processing is Executed\n");
            #endregion

            #region MyRegion
            //Console.WriteLine("\n stored procedure procProcessEmployeeUpdates started\n");
            //string Query9 = "exec procProcessEmployeeUpdates";
            //SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            //while (datareader9.Read())
            //{
            //    Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            //}
            //Console.WriteLine("\nstored procedure procProcessEmployeeUpdates is Executed\n");
            #endregion
        }
    }
}
