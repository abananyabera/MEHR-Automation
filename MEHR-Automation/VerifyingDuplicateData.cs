using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Policy;


namespace MEHR_Automation
{
    public class VerifyingDuplicateData
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        Update_duplicates update_Duplicates = new Update_duplicates();

        public void VerifyDuplicateDataintable(SqlConnection sqlconnection)
        {
            
            Console.WriteLine("Duplicates verfication is started");
            
            string Query = "select uniqueid,count(uniqueid),datasourceid from tbl_employees_import\r\ngroup by uniqueid,datasourceid\r\nhaving count(uniqueid)>1\r\n\r\nIF OBJECT_ID('tbl_Employees_Import_Excluded') IS NOT NULL \r\n\tDROP TABLE tbl_Employees_Import_Excluded\r\n\t\r\n\r\nselect b.*\r\ninto tbl_Employees_Import_Excluded\r\nfrom WorkdayIntegratedEmployees a, tbl_employees_import b\r\nwhere me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid\r\n\r\nDELETE b\r\nfrom WorkdayIntegratedEmployees a, tbl_employees_import b\r\nwhere me_uniqueID=b.uniqueid and a.datasourceid=b.datasourceid;";


            using (SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection))
            {
                while (datareader.Read())
                {
                    if (string.IsNullOrEmpty(datareader[0].ToString()))
                    {
                        Console.WriteLine("Returned value is null.");
                    }
                    else
                    {
                        Console.WriteLine("Not Returning Empty results in current Executing Query");
                        Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
                        //Environment.Exit(0);
                    }
                   
                }
                Console.WriteLine("Duplicates verfication is successfull ");
            }

        }

        public void reverifyDuplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("Duplicates reverfication is started");
            string Query = "Select epassid from tbl_employees_stage1 where active = 1 group by epassid having count(epassid)>1";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (!datareader.HasRows)
            {
                while (datareader.Read())
                {
                    Console.WriteLine(datareader[0]);
                }
                Console.WriteLine("Duplicates reverfication is successfull");
                ReadLine();
            }
            else
            {
                while (datareader.Read())
                {
                    if (string.IsNullOrEmpty(datareader[0].ToString()))
                    {
                        Console.WriteLine("Returned value is null.");
                    }
                    else
                    {
                        int count = 0;
                        var Temp_old_masterid = "";
                        var Temp_new_masterid = "";
                        string select_Query = "SELECT * FROM tbl_employees_stage1 WHERE epassid = " + datareader[0];
                        Console.WriteLine(select_Query);
                        SqlDataReader selectQuery_datareader = executeQueries.ExecuteQuery(select_Query, sqlconnection);
                        Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}", selectQuery_datareader.GetName(0), selectQuery_datareader.GetName(3), selectQuery_datareader.GetName(4), selectQuery_datareader.GetName(5), selectQuery_datareader.GetName(6), selectQuery_datareader.GetName(7), selectQuery_datareader.GetName(8), selectQuery_datareader.GetName(23), selectQuery_datareader.GetName(24), selectQuery_datareader.GetName(28), selectQuery_datareader.GetName(29));
                        while (selectQuery_datareader.Read())
                        {
                            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}",
                            selectQuery_datareader[0], selectQuery_datareader[3], selectQuery_datareader[4],
                            selectQuery_datareader[5], selectQuery_datareader[6], selectQuery_datareader[7],
                            selectQuery_datareader[8], selectQuery_datareader[23], selectQuery_datareader[24],
                            selectQuery_datareader[28], selectQuery_datareader[29]); 
                            if (count == 0)
                            {
                                Temp_old_masterid = Convert.ToString(selectQuery_datareader[0]);
                            }
                            if (count == 1)
                            {
                                Temp_new_masterid = Convert.ToString(selectQuery_datareader[0]);
                                update_Duplicates.Updating_duplicate(Temp_new_masterid, Temp_old_masterid, sqlconnection);
                            }
                            count++;

                        }
                    }

                }
            }


            Console.WriteLine("\n Duplicates reverfication is started");
            string Query2 = "Select internet_email from tbl_employees_stage1 where active = 1 group by internet_email having count(internet_email)>1";
            SqlDataReader datareader2 = executeQueries.ExecuteQuery(Query2, sqlconnection);
            if (!datareader2.HasRows)
            {
                while (datareader2.Read())
                {
                    Console.WriteLine(datareader2[0]);
                }
                Console.WriteLine("\n Duplicates reverfication is successfull");
                ReadLine();
            }
            else
            {
                while (datareader2.Read())
                {
                    if (string.IsNullOrEmpty(datareader2[0].ToString()))
                    {
                        Console.WriteLine("Returned value is null.");
                    }
                    else
                    {
                        int count = 0;
                        var Temp_old_masterid = "";
                        var Temp_new_masterid = "";
                        string select_Query = "SELECT * FROM tbl_employees_stage1 WHERE epassid = " + datareader2[0];
                        Console.WriteLine(select_Query);
                        SqlDataReader selectQuery_datareader = executeQueries.ExecuteQuery(select_Query, sqlconnection);
                        Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}", selectQuery_datareader.GetName(0), selectQuery_datareader.GetName(3), selectQuery_datareader.GetName(4), selectQuery_datareader.GetName(5), selectQuery_datareader.GetName(6), selectQuery_datareader.GetName(7), selectQuery_datareader.GetName(8), selectQuery_datareader.GetName(23), selectQuery_datareader.GetName(24), selectQuery_datareader.GetName(28), selectQuery_datareader.GetName(29));
                        while (selectQuery_datareader.Read())
                        {
                            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}",
                            selectQuery_datareader[0], selectQuery_datareader[3], selectQuery_datareader[4],
                            selectQuery_datareader[5], selectQuery_datareader[6], selectQuery_datareader[7],
                            selectQuery_datareader[8], selectQuery_datareader[23], selectQuery_datareader[24],
                            selectQuery_datareader[28], selectQuery_datareader[29]);
                            if (count == 0)
                            {
                                Temp_old_masterid = Convert.ToString(selectQuery_datareader[0]);
                            }
                            if (count == 1)
                            {
                                Temp_new_masterid = Convert.ToString(selectQuery_datareader[0]);
                                update_Duplicates.Updating_duplicate(Temp_new_masterid, Temp_old_masterid, sqlconnection);
                            }
                            count++;

                        }
                    }

                }
            }


            Console.WriteLine("Duplicates reverfication is started");
            string Query3 = "Select epassid from tbl_employees_stage1 group by epassid having count(epassid)>1";
            SqlDataReader datareader3 = executeQueries.ExecuteQuery(Query3, sqlconnection);
            if (!datareader3.HasRows)
            {
                while (datareader3.Read())
                {
                    Console.WriteLine(datareader3[0]);
                }
                Console.WriteLine("\n Duplicates reverfication is successfull");
                ReadLine();
            }
            else
            {
                while (datareader3.Read())
                {
                    if (string.IsNullOrEmpty(datareader3[0].ToString()))
                    {
                        Console.WriteLine("Returned value is null.");
                    }
                    else
                    {
                        int count = 0;
                        var Temp_old_masterid = "";
                        var Temp_new_masterid = "";
                        string select_Query = "SELECT * FROM tbl_employees_stage1 WHERE epassid = " + datareader3[0];
                        Console.WriteLine(select_Query);
                        SqlDataReader selectQuery_datareader = executeQueries.ExecuteQuery(select_Query, sqlconnection);
                        Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}", selectQuery_datareader.GetName(0), selectQuery_datareader.GetName(3), selectQuery_datareader.GetName(4), selectQuery_datareader.GetName(5), selectQuery_datareader.GetName(6), selectQuery_datareader.GetName(7), selectQuery_datareader.GetName(8), selectQuery_datareader.GetName(23), selectQuery_datareader.GetName(24), selectQuery_datareader.GetName(28), selectQuery_datareader.GetName(29));
                        while (selectQuery_datareader.Read())
                        {
                            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}",
                            selectQuery_datareader[0], selectQuery_datareader[3], selectQuery_datareader[4],
                            selectQuery_datareader[5], selectQuery_datareader[6], selectQuery_datareader[7],
                            selectQuery_datareader[8], selectQuery_datareader[23], selectQuery_datareader[24],
                            selectQuery_datareader[28], selectQuery_datareader[29]);
                            if (count == 0)
                            {
                                Temp_old_masterid = Convert.ToString(selectQuery_datareader[0]);
                            }
                            if (count == 1)
                            {
                                Temp_new_masterid = Convert.ToString(selectQuery_datareader[0]);
                                update_Duplicates.Updating_duplicate(Temp_new_masterid, Temp_old_masterid, sqlconnection);
                            }
                            count++;

                        }
                    }
                    
                }


            }
        }
        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }

    }
}
