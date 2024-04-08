using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using System.Configuration;

namespace MEHR_Automation
{
    public class StoredProcedure
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        Update_duplicates update_Duplicates = new Update_duplicates();

        public void StoredProcedureExecution(SqlConnection sqlconnection)
        {

            Console.WriteLine("\nstored procedure proc_Pre_Update_Processing started");
            string proc_Pre_Update_Processing_Query = "exec proc_Pre_Update_Processing";
            SqlDataReader datareader = executeQueries.ExecuteQuery(proc_Pre_Update_Processing_Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1]);
            }
            Console.WriteLine("\nstored procedure proc_Pre_Update_Processing is Executed");


            Console.WriteLine("\n stored procedure procProcessEmployeeUpdates started");
            string procProcessEmployeeUpdates_Query = "exec procProcessEmployeeUpdates";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(procProcessEmployeeUpdates_Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procProcessEmployeeUpdates is Executed");
        }

        public void procUpdatetempfixStoredProcedure(SqlConnection sqlconnection)
        {

            Console.WriteLine("\n stored procedure procUpdateSolaeJobSubgroup_tempfix started");
            string Query = "exec procUpdateSolaeJobSubgroup_tempfix";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdateSolaeJobSubgroup_tempfix is Executed");
        }

        public void procPostUpdateStoredProcedure(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure proc_Post_Update_Processing started");
            string Query = "exec proc_Post_Update_Processing";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if (dataReader.HasRows)
            {
                Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15}", dataReader.GetName(0), dataReader.GetName(1), dataReader.GetName(2), dataReader.GetName(3), dataReader.GetName(4), dataReader.GetName(5), dataReader.GetName(6));
                while (dataReader.Read())
                {
                    int count = 0;
                    var Temp = "";
                    var Temp1 = "";
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15}", dataReader[0], dataReader[1], dataReader[2], dataReader[3], dataReader[4], dataReader[5], dataReader[6]);
                    var newmasterid = Convert.ToString(dataReader[5]);
                    var oldmasterid = Convert.ToString(dataReader[6]);
                    string selectQuery = "select * from tbl_employees_stage1 where masterid in ("+ newmasterid +"," + oldmasterid + ")";
                    SqlDataReader selectQuerydatareader = executeQueries.ExecuteQuery(selectQuery, sqlconnection);
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}", selectQuerydatareader.GetName(0), selectQuerydatareader.GetName(3), selectQuerydatareader.GetName(4), selectQuerydatareader.GetName(5), selectQuerydatareader.GetName(6), selectQuerydatareader.GetName(7), selectQuerydatareader.GetName(8), selectQuerydatareader.GetName(23), selectQuerydatareader.GetName(24), selectQuerydatareader.GetName(28), selectQuerydatareader.GetName(29));
                    while (selectQuerydatareader.Read())
                    {
                       
                        Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15} | {9,-15} | {10,-15}",
                        selectQuerydatareader[0], selectQuerydatareader[3], selectQuerydatareader[4],
                        selectQuerydatareader[5], selectQuerydatareader[6], selectQuerydatareader[7],
                        selectQuerydatareader[8], selectQuerydatareader[23], selectQuerydatareader[24], 
                        selectQuerydatareader[28], selectQuerydatareader[29]);
                        Console.WriteLine("\n Plese check the  above table Mainly the Epass id ");
                        Console.WriteLine("\n Note: If the Epass id are different then the both the Users are Different. Suppose if it is a duplicate then we need to handle the duplicate.");
                        Console.WriteLine("we have to handle the duplicate in by making one as primary and other as secondary");
                        Console.WriteLine("Please click Enter to procedd");
                        ReadLine();
                        if (count == 0)
                        {
                            Temp = Convert.ToString(selectQuerydatareader[28]);
                            Temp1 = Convert.ToString(selectQuerydatareader[29]);
                        }
                        if(count == 1)
                        {
                            var old_master_id = Convert.ToString(selectQuerydatareader[28]);
                            var new_master_id = Convert.ToString(selectQuerydatareader[29]);
                            if (old_master_id == Temp || new_master_id == Temp1)
                            {
                                Console.WriteLine("\nThe epass id are matching we have the action item i.e. new record is the primary and old record is the secondary");
                                Console.WriteLine("\nNext action : Please click enter to verfy in the duplicated first if recored is present then there is no action else there is action item");
                                
                                //update query
                                updateduplicates(newmasterid, oldmasterid, sqlconnection);
                                Console.WriteLine("\n The newmasterid is marked as primary kindly validate the same.  please click enter to mark the secondary as duplicate");
                                ReadLine();

                                //update the secondary master id as null
                                update_Duplicates.Secondary_Masterid_As_Null(newmasterid, oldmasterid, sqlconnection);

                                //For primary set the email type id = 1
                                update_Duplicates.set_primary_Email_type_id(newmasterid, oldmasterid, sqlconnection);

                                //Updating and reverifying duplicates
                                update_Duplicates.procUpdateDuplicateInformation(sqlconnection);

                                update_Duplicates.ReverifyingDuplicates(newmasterid, oldmasterid, sqlconnection);
                            }
                            else
                            {
                                Console.WriteLine(" The Epass Id are different so there are two different users hence there is no action item from our end");
                            }
                        }
                        count++;
                    }
                    

                }
            }
            else
            {
                Console.WriteLine(" ** All is Good we can move to the next steps ** ");
            }
            Console.WriteLine("stored procedure proc_Post_Update_Processing is Executed");

        }

        public void updateduplicates(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            string varify_duplicate = "select * from tbl_duplicates where PrimaryMasterid in (" + newmasterid + "," + oldmasterid + ") and SecondaryMasterID in (" + newmasterid + "," + oldmasterid + ")";
            SqlDataReader duplicate_datareader = executeQueries.ExecuteQuery(varify_duplicate, sqlconnection);
            if (!duplicate_datareader.HasRows)
            {
                while (duplicate_datareader.Read())
                {
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader.GetName(0), duplicate_datareader.GetName(1), duplicate_datareader.GetName(2), duplicate_datareader.GetName(3), duplicate_datareader.GetName(4), duplicate_datareader.GetName(5));
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader[0], duplicate_datareader[1], duplicate_datareader[2], duplicate_datareader[3], duplicate_datareader[4], duplicate_datareader[5]);

                    //insert duplicates
                    update_Duplicates.insertduplidate(newmasterid, oldmasterid, sqlconnection);

                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader.GetName(0), duplicate_datareader.GetName(1), duplicate_datareader.GetName(2), duplicate_datareader.GetName(3), duplicate_datareader.GetName(4), duplicate_datareader.GetName(5));
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader[0], duplicate_datareader[1], duplicate_datareader[2], duplicate_datareader[3], duplicate_datareader[4], duplicate_datareader[5]);
                }

            }
            else
            {
                Console.WriteLine(" Data is present iin the 'tbl_duplicate' then No Action From oue End.");
            }
        }

   
        

        public void procRemoveInvalidEmail_SP(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRemoveInvalidEmail started");
            string Query = "exec dbo.procRemoveInvalidEmail";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procRemoveInvalidEmail is Executed");

        }
        
        

        public void procRefineManager_SP(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procRefineManagerData started");
            string Query = "exec procRefineManagerData";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procRefineManagerData is Executed");

        }

        public void procOverrideBadUpdates_SP(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procOverrideBadUpdates started");
            string Query = "exec dbo.procOverrideBadUpdates";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procOverrideBadUpdates is Executed");

        }

        public void procStage1ReviewCleanup_SP(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procStage1ReviewCleanup started");
            string Query = "exec procStage1ReviewCleanup";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procStage1ReviewCleanup is Executed");

        }

        public void procUpdatePicklistValues_SP(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procUpdatePicklistValues started");
            string Query = "exec procUpdatePicklistValues";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdatePicklistValues is Executed");

        }

        public void replaceSpecialChar(SqlConnection sqlconnection)
        {
            Console.WriteLine("\nstored procedure replaceSpecialChar started");
            string Query = "exec Proc_ReplaceSpecialChar";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure replaceSpecialChar is Executed");

        }

        


        //Will execute al last after finishing everything
        public void procUpdateMasterEmployeeTable(SqlConnection sqlconnection)
        {
            Console.WriteLine(" stored procedure procUpdateMasterEmployeeTable started ");
            string Query = "exec procUpdateMasterEmployeeTable";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("stored procedure procUpdateMasterEmployeeTable is Executed");

        }

        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }


    }
}
