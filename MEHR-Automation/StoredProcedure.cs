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
            //string proc_Pre_Update_Processing_Query = "exec proc_Pre_Update_Processing";
            //SqlDataReader datareader = executeQueries.ExecuteQuery(proc_Pre_Update_Processing_Query, sqlconnection);
            //while (datareader.Read())
            //{
            //    Console.WriteLine(datareader[0] + "|" + datareader[1]);
            //}
            //Console.WriteLine("\nstored procedure proc_Pre_Update_Processing is Executed\n");
            #endregion

            #region MyRegion
            //Console.WriteLine("\n stored procedure procProcessEmployeeUpdates started\n");
            //string procProcessEmployeeUpdates_Query = "exec procProcessEmployeeUpdates";
            //SqlDataReader datareader9 = executeQueries.ExecuteQuery(procProcessEmployeeUpdates_Query, sqlconnection);
            //while (datareader9.Read())
            //{
            //    Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            //}
            //Console.WriteLine("\nstored procedure procProcessEmployeeUpdates is Executed\n");

            #endregion
        }

        public void procUpdatetempfixStoredProcedure(SqlConnection sqlconnection)
        {

            Console.WriteLine("\n stored procedure procUpdateSolaeJobSubgroup_tempfix started\n");
            string Query9 = "exec procUpdateSolaeJobSubgroup_tempfix";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procUpdateSolaeJobSubgroup_tempfix is Executed\n");
        }

        public void procPostUpdateStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure proc_Post_Update_Processing started\n");
            string Query9 = "exec proc_Post_Update_Processing";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure proc_Post_Update_Processing is Executed\n");

        }

        public void procRemoveInvalidEmailStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRemoveInvalidEmail started\n");
            string Query9 = "exec dbo.procRemoveInvalidEmail";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procRemoveInvalidEmail is Executed\n");

        }

        public void procRefineManagerDataStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRefineManagerData started\n");
            string Query9 = "exec procRefineManagerData";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procRefineManagerData is Executed\n");

        }

        public void procOverrideBadUpdatesStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procOverrideBadUpdates started\n");
            string Query9 = "exec dbo.procOverrideBadUpdates";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procOverrideBadUpdates is Executed\n");

        }

        public void procStage1ReviewCleanupStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procStage1ReviewCleanup started\n");
            string Query9 = "exec procStage1ReviewCleanup";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procStage1ReviewCleanup is Executed\n");

        }

        public void procUpdatePicklistValuesStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procUpdatePicklistValues started\n");
            string Query9 = "exec procUpdatePicklistValues";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\nstored procedure procUpdatePicklistValues is Executed\n");

        }


    }
}
