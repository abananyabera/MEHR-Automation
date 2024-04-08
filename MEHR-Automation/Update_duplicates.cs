using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace MEHR_Automation
{
    public class Update_duplicates
    {
        ExecuteQueries executeQueries = new ExecuteQueries();

        public void Updating_duplicate(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            insertduplidate(newmasterid, oldmasterid, sqlconnection);
            Console.WriteLine("\n The newmasterid is marked as primary kindly validate the same.  please click enter to mark the secondary as duplicate");
            ReadLine();

            //update the secondary master id as null
            Secondary_Masterid_As_Null(newmasterid, oldmasterid, sqlconnection);

            //For primary set the email type id = 1
            set_primary_Email_type_id(newmasterid, oldmasterid, sqlconnection);

            //Updating and reverifying duplicates
            procUpdateDuplicateInformation(sqlconnection);
            ReverifyingDuplicates(newmasterid, oldmasterid, sqlconnection);
        }
        public void insertduplidate(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            string insertQuery_duplicates = "insert into tbl_duplicates values (" + newmasterid + "," + oldmasterid + ",0,0,getdate()";
            SqlDataReader insertQuery_duplicate_datareader = executeQueries.ExecuteQuery(insertQuery_duplicates, sqlconnection);
        }
        public void Secondary_Masterid_As_Null(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            string Making_secondary_Masterid_NULL = "update tbl_employees_stage1 set epassid = '', email_type_id = NULL where masterid = " + oldmasterid;
            SqlDataReader update_secondary_masterid_datareader = executeQueries.ExecuteQuery(Making_secondary_Masterid_NULL, sqlconnection);

        }

        public void set_primary_Email_type_id(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            string Updating_primary_emailtypeid = "update tbl_employees_stage1 set email_type_id = NULL where masterid = " + newmasterid;
            SqlDataReader primary_emailtypeid_datareader = executeQueries.ExecuteQuery(Updating_primary_emailtypeid, sqlconnection);

        }
        public void ReverifyingDuplicates(string newmasterid, string oldmasterid, SqlConnection sqlconnection)
        {
            string Revarify_duplicate = "select * from tbl_duplicates where PrimaryMasterid in (" + newmasterid + "," + oldmasterid + ") and SecondaryMasterID in (" + newmasterid + "," + oldmasterid + ")";
            SqlDataReader duplicate_datareader = executeQueries.ExecuteQuery(Revarify_duplicate, sqlconnection);
            while (duplicate_datareader.Read())
            {
                Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader.GetName(0), duplicate_datareader.GetName(1), duplicate_datareader.GetName(2), duplicate_datareader.GetName(3), duplicate_datareader.GetName(4), duplicate_datareader.GetName(5));
                Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15}", duplicate_datareader[0], duplicate_datareader[1], duplicate_datareader[2], duplicate_datareader[3], duplicate_datareader[4], duplicate_datareader[5]);

            }
        }

        public void procUpdateDuplicateInformation(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n stored procedure procUpdateDuplicateInformation started");
            string Query = "exec procUpdateDuplicateInformation";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (dataReader.Read())
            {
                Console.WriteLine(dataReader[0] + "|" + dataReader[1]);
            }
            Console.WriteLine("\nstored procedure procUpdateDuplicateInformation is Executed");

        }

        

        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }
    }

}
