﻿using System.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class tablebackup
    {

        ExecuteQueries executeQueries = new ExecuteQueries();
       

        public void takeTableBackup(SqlConnection sqlconnection) {
            string timeStamp = DateTime.Now.ToString("MMddyyyy");
            string destinationTable1 = "[dbo]. [tbl_employees_stage1_" + timeStamp + "]";
            string query1 = "select * into" + " " + destinationTable1 + " " + "from [dbo]. [tbl_employees_stage1]";


            // checking the table is already presnet or not if present returns 1 else return 0
            int connectionresult = 0;
            string checkingtable = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_" + timeStamp + "'";
            SqlDataReader connection = executeQueries.ExecuteQuery(checkingtable, sqlconnection);
            while (connection.Read())
            {
                connectionresult = (int)connection[0];
            }

            //drop backup table if already present
            if (connectionresult == 1)
            {
                SqlCommand cmd = new SqlCommand("drop table " + destinationTable1, sqlconnection);
                cmd.ExecuteNonQuery();
                Console.WriteLine(destinationTable1 + "dropped succesfully");
            }

            // creates the backuptable 1 if not present else it returns error
            executeQueries.ExecuteQuery(query1, sqlconnection);


            int countMainTable = 0;
            string query4 = "Select count(*) from [dbo]. [tbl_employees_stage1]";
            SqlDataReader counter0 = executeQueries.ExecuteQuery(query4, sqlconnection);
            while (counter0.Read())
            {
                countMainTable = (int)counter0[0];
            }

            int countMainTableBackup = 0;
            string query5 = "Select count(*) from " + destinationTable1;
            SqlDataReader counter1 = executeQueries.ExecuteQuery(query5, sqlconnection);
            while (counter1.Read())
            {
                countMainTableBackup = (int)counter1[0];
            }

            if (countMainTable == countMainTableBackup)
            {
                Console.WriteLine("Backup for tbl_employees_stage1 is successful");
            }
            else
            {
                Console.WriteLine("Backup for tbl_employees_stage1 is failed");
            }


            string timeStamp2 = DateTime.Now.ToString("MMddyyyy");
            string destinationTable2 = "[dbo]. [tbl_employees_stage1_hold_" + timeStamp2 + "]";
            string backupquery = "select * into" + " " + destinationTable2 + " " + "from [dbo]. [tbl_employees_stage1_hold]";


            // checking the table is already presnet or not if present returns 1 else return 0
            int connectionresult2 = 0;
            string checkingtable2 = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_hold_" + timeStamp2 + "'";

            SqlDataReader connection2 = executeQueries.ExecuteQuery(checkingtable2, sqlconnection);
            while (connection2.Read())
            {
                connectionresult2 = (int)connection2[0];

            }

            //drop backup table 2 if already present
            if (connectionresult2 == 1)
            {
                SqlCommand cmd = new SqlCommand("drop table " + destinationTable2, sqlconnection);

                cmd.ExecuteNonQuery();
                Console.WriteLine(destinationTable2 + "dropped succesfully");
            }

            // creates the backuptable 2 if not present else it returns error
            executeQueries.ExecuteQuery(backupquery, sqlconnection);


            int countMainTable2 = 0;
            string backupquery4 = "Select count(*) from tbl_employees_stage1_hold";
            SqlDataReader backupcounter1 = executeQueries.ExecuteQuery(backupquery4, sqlconnection);
            while (counter0.Read())
            {
                countMainTable2 = (int)backupcounter1[0];
            }

            int countMainTableBackup2 = 0;
            string backquery5 = "Select count(*) from " + destinationTable2;
            SqlDataReader backupcounter2 = executeQueries.ExecuteQuery(backquery5, sqlconnection);
            while (counter1.Read())
            {
                countMainTableBackup2 = (int)backupcounter2[0];
            }

            if (countMainTable2 == countMainTableBackup2)
            {
                Console.WriteLine("Backup for tbl_employees_stage1_hold is successful");
            }
            else
            {
                Console.WriteLine("Backup for tbl_employees_stage1_hold is failed");
            }
        }
    }
}
