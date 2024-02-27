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
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, upper([corporate_entity]) AS Expr1, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.entityid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY Upper([corporate_entity]), dbo.tbl_Employees_Stage1.datasourceid";
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
                }
            }
        }

        public void UnmappedSBUs(SqlConnection sqlconnection)
        {
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.sbu, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.sbuid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY dbo.tbl_Employees_Stage1.sbu, dbo.tbl_Employees_Stage1.datasourceid";
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
                }
            }
        }

        public void UnmappedSites(SqlConnection sqlconnection)
        {
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
                }
            }
        }

        public void UnmappedOps(SqlConnection sqlconnection)
        {
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.platform, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.platformid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY dbo.tbl_Employees_Stage1.platform, dbo.tbl_Employees_Stage1.datasourceid";
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
                }
            }
        }

        public void UnmappedFunctions(SqlConnection sqlconnection)
        {
            string Query = "SELECT DISTINCT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, upper([busfunction]) AS Expr1, dbo.tbl_Employees_Stage1.datasourceid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.functionid) Is Null) AND ((dbo.tbl_Employees_Stage1.active)=1))\r\nGROUP BY Upper([busfunction]), dbo.tbl_Employees_Stage1.datasourceid;";
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
                }
            }
        }





    }
}
