using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using static System.Net.Mime.MediaTypeNames;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.Windows.Forms;


namespace MEHR_Automation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            sqlconnection.Open();
            Console.WriteLine("Connection is successfull");
            //execute commands
            Console.WriteLine("--------------------------------------------------------------------------------------------");
            string timeStamp = DateTime.Now.ToString("MMddyyyy");
            string destinationTable1 = "[dbo]. [tbl_employees_stage1_" + timeStamp+ "]";
            string query1 = "select * into"+" "+ destinationTable1+" "+"from [dbo]. [tbl_employees_stage1]";
            

            // checking the table is already presnet or not if present returns 1 else return 0
            int connectionresult = 0;
            string checkingtable = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_" + timeStamp  + "'";
            SqlDataReader connection = ExecuteQuery(checkingtable, sqlconnection);
            while (connection.Read())
            {
                connectionresult = (int)connection[0];
            }

            //drop backup table if already present
            if (connectionresult == 1)
            {
                SqlCommand cmd = new SqlCommand("drop table [dbo]. [tbl_employees_stage1_01252024]", sqlconnection);
                cmd.ExecuteNonQuery();
                Console.WriteLine(destinationTable1 + "dropped succesfully");
            }

            // creates the backuptable 1 if not present else it returns error
            ExecuteQuery(query1, sqlconnection);


            int countMainTable = 0;
            string query4 = "Select count(*) from tbl_employees_stage1";
            SqlDataReader counter0 = ExecuteQuery(query4, sqlconnection);
            while (counter0.Read())
            {
                countMainTable = (int)counter0[0];
            }

            int countMainTableBackup = 0;
            string query5 = "Select count(*) from [dbo]. [tbl_employees_stage1_01252024]";
            SqlDataReader counter1 = ExecuteQuery(query5, sqlconnection);
            while (counter1.Read())
            {
                countMainTableBackup = (int)counter1[0];
            }

            if(countMainTable == countMainTableBackup)
            {
                Console.WriteLine("Backup for tbl_employees_stage1 is successful");
            }
            else
            {
                Console.WriteLine("Backup for tbl_employees_stage1 is failed");
            }
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            string query2 = "Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid";
            SqlDataReader datareader = ExecuteQuery(query2, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0]+ "|" + datareader[1] + "|" + datareader[2]);
            }

            sqlconnection.Close();
            Console.ReadLine();
        }

        public static SqlDataReader ExecuteQuery(string query, SqlConnection connection) {
            try
            {
                SqlCommand cmd = new SqlCommand(query, connection);
                SqlDataReader datareader = cmd.ExecuteReader();
                return datareader;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error executing query: {ex.Message}");
                throw;
            }
        }
    }
}
