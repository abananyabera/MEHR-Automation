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
            //SqlDataReader dataReader = executeQueries.ExecuteQuery(procProcessEmployeeUpdates_Query, sqlconnection);
            //while (dataReader.Read())
            //{
            //    Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            //}
            //Console.WriteLine("\nstored procedure procProcessEmployeeUpdates is Executed\n");

            #endregion
        }

        public void procUpdatetempfixStoredProcedure(SqlConnection sqlconnection)
        {

            Console.WriteLine("\n stored procedure procUpdateSolaeJobSubgroup_tempfix started\n");
            string Query = "exec procUpdateSolaeJobSubgroup_tempfix";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdateSolaeJobSubgroup_tempfix is Executed\n");
        }

        public void procPostUpdateStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure proc_Post_Update_Processing started\n");
            string Query = "exec proc_Post_Update_Processing";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure proc_Post_Update_Processing is Executed\n");

        }

        public void procRemoveInvalidEmailStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRemoveInvalidEmail started\n");
            string Query = "exec dbo.procRemoveInvalidEmail";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procRemoveInvalidEmail is Executed\n");

        }

        public void procRefineManagerDataStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRefineManagerData started\n");
            string Query = "exec procRefineManagerData";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procRefineManagerData is Executed\n");

        }

        public void procOverrideBadUpdatesStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procOverrideBadUpdates started\n");
            string Query = "exec dbo.procOverrideBadUpdates";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procOverrideBadUpdates is Executed\n");

        }

        public void procStage1ReviewCleanupStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procStage1ReviewCleanup started\n");
            string Query = "exec procStage1ReviewCleanup";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procStage1ReviewCleanup is Executed\n");

        }

        public void procUpdatePicklistValuesStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procUpdatePicklistValues started\n");
            string Query = "exec procUpdatePicklistValues";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdatePicklistValues is Executed\n");

        }

        public void replaceSpecialChar(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure replaceSpecialChar started\n");
            string Query = "exec Proc_ReplaceSpecialChar";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure replaceSpecialChar is Executed\n");

        }

        public void findSpecialChar(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure findSpeciaChar started\n");
            string Query = "exec find_Specialchar";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}",dataReader.GetName(0), dataReader.GetName(1), dataReader.GetName(2), dataReader.GetName(3));
            while (dataReader.Read())
            {
                Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader[0], dataReader[1], dataReader[2], dataReader[3]);

            }
            Console.WriteLine("\nstored procedure findSpecialChar is Executed\n");

        }

        //Will execute al last after finishing everything
        public void procUpdateMasterEmployeeTable(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procUpdateMasterEmployeeTable started\n");
            string Query = "exec procUpdateMasterEmployeeTable";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdateMasterEmployeeTable is Executed\n");

        }


    }
}
