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
            Console.WriteLine("\n Duplicates verfication is started\n");
            string Query = "select uniqueid,count(uniqueid),datasourceid from tbl_employees_import\r\ngroup by uniqueid,datasourceid\r\nhaving count(uniqueid)>1\r\n\r\nIF OBJECT_ID('tbl_Employees_Import_Excluded') IS NOT NULL \r\n\tDROP TABLE tbl_Employees_Import_Excluded\r\n\t\r\n\r\nselect b.*\r\ninto tbl_Employees_Import_Excluded\r\nfrom WorkdayIntegratedEmployees a, tbl_employees_import b\r\nwhere me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid\r\n\r\nDELETE b\r\nfrom WorkdayIntegratedEmployees a, tbl_employees_import b\r\nwhere me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (datareader.HasRows)
            {
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
                }
                Console.WriteLine("\n Duplicates verfication is successfull \n");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

                //Environment.Exit(0);
            }
            #endregion


        }

        public void reverifyDuplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n Duplicates reverfication is started\n");
            string Query = "Select epassid from tbl_employees_stage1 where active = 1 group by epassid having count(epassid)>1";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (datareader.HasRows)
            {
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0]);
                }
                Console.WriteLine("\n Duplicates reverfication is successfull \n");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

            }


            Console.WriteLine("\n Duplicates reverfication is started\n");
            string Query2 = "Select internet_email from tbl_employees_stage1 where active = 1 group by internet_email having count(internet_email)>1";
            SqlDataReader datareader2 = executeQueries.ExecuteQuery(Query2, sqlconnection);
            if (datareader2.HasRows)
            {
                while (datareader2.Read())
                {
                    Console.WriteLine(datareader2[0] );
                }
                Console.WriteLine("\n Duplicates reverfication is successfull \n");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

            }


            Console.WriteLine("\n Duplicates reverfication is started\n");
            string Query3 = "Select epassid from tbl_employees_stage1 group by epassid having count(epassid)>1";
            SqlDataReader datareader3 = executeQueries.ExecuteQuery(Query3, sqlconnection);
            if (datareader3.HasRows)
            {
                while (datareader3.Read())
                {
                    Console.WriteLine(datareader3[0]);
                }
                Console.WriteLine("\n Duplicates reverfication is successfull \n");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();
            }
            else
            {
                Console.WriteLine(" Not Returning Empty results in current Executing Query");
                Console.WriteLine("We cannot proceed any further please type any key to Exit");
                Console.WriteLine("-------------------------------------------------------------");
                Console.ReadLine();

            }
        }

    }
}
