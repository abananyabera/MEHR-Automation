using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class PicklingQueries
    {
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void UnMappedEntities(SqlConnection sqlconnection)
        {
            Console.WriteLine("The UnMappedEntities pickling query started Execution");
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, upper([corporate_entity]) AS Expr1, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.entityid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY Upper([corporate_entity]), dbo.tbl_Employees_Stage1.datasourceid";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            StoredProcedure storedProcedure = new StoredProcedure();
            if (!datareader.HasRows)
            {
                Console.WriteLine("No results Found");
            }
            else
            {
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2]  );
                    string selectQuery = "select * from tbl_entity where description like '%" + datareader[1] +"%'";
                    SqlDataReader queryReturned = executeQueries.ExecuteQuery(selectQuery, sqlconnection);
                    if (!queryReturned.HasRows)
                    {
                        string selectSrNoQuery = "select entityid from tbl_Entity where description = '" + queryReturned[2] + "'";
                        SqlDataReader srNo = executeQueries.ExecuteQuery(selectSrNoQuery, sqlconnection);

                        string insertQuery = "insert into tbl_entity values ('" + srNo[0] + "','" + queryReturned[2] +"',1,1)";
                        SqlDataReader Insert = executeQueries.ExecuteQuery(insertQuery, sqlconnection);
                        storedProcedure.procUpdatePicklistValues_SP(sqlconnection);
                    }
                }
            }
        }

        public void UnmappedSBUs(SqlConnection sqlconnection)
        {
            Console.WriteLine("The UnmappedSBUs pickling query started Execution");
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.sbu, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.sbuid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY dbo.tbl_Employees_Stage1.sbu, dbo.tbl_Employees_Stage1.datasourceid";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (!datareader.HasRows)
            {
                Console.WriteLine("No results Found");
            }
            else
            {
                Console.WriteLine(" We are having the records in the 'UnmappedSBUs' please validate Manually");
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] );
                }
            }
        }

        public void UnmappedSites(SqlConnection sqlconnection)
        {
            Console.WriteLine("The UnmappedSites pickling query started Execution");
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, upper([site]) AS Expr1, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Employees_Stage1.country, dbo.tbl_Employees_Stage1.state\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.siteid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY upper([site]), dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Employees_Stage1.country, dbo.tbl_Employees_Stage1.state;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (!datareader.HasRows)
            {
                Console.WriteLine("No results Found");
            }
            else
            {
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2]  );
                    string query = "select * from tbl_Site where description like '%" + datareader[1] +"%'";
                    SqlDataReader queryReturned = executeQueries.ExecuteQuery(query, sqlconnection);
                    if (!queryReturned.HasRows)
                    {
                        string queryCountryId = "select * from tbl_Country where description like '%" + datareader[3] +"%'";
                        SqlDataReader returned = executeQueries.ExecuteQuery(query, sqlconnection);
                        string queryInsert = "insert into tbl_Site VALUES (" + returned[0] + "," + returned[1] +" , 0, 'ABC', '" + returned[3] +"',1,1)";
                        SqlDataReader insertData = executeQueries.ExecuteQuery(queryInsert, sqlconnection);

                    }

                }
            }
        }

        public void UnmappedOps(SqlConnection sqlconnection)
        {
            Console.WriteLine("The UnmappedOps pickling query started Execution");
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.platform, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.platformid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY dbo.tbl_Employees_Stage1.platform, dbo.tbl_Employees_Stage1.datasourceid";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (!datareader.HasRows)
            {
                Console.WriteLine("No results Found");
            }
            else
            {
                Console.WriteLine(" We are having the records in the 'UnmappedOps' please validate Manually");
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2]  );
                }
            }
        }

        public void UnmappedFunctions(SqlConnection sqlconnection)
        {
            Console.WriteLine("The UnmappedFunctions pickling query started Execution");
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, upper([busfunction]) AS Expr1, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.functionid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY Upper([busfunction]), dbo.tbl_Employees_Stage1.datasourceid;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (!datareader.HasRows)
            {
                Console.WriteLine("No results Found");
            }
            else
            {
                Console.WriteLine(" We are having the records in the 'UnmappedFunctions' please validate Manually");
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2]  );
                }
            }
        }





    }
}
