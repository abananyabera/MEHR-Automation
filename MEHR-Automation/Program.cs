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
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.IO;


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

            #region MyRegion

            string timeStamp = DateTime.Now.ToString("MMddyyyy");
            string destinationTable1 = "[dbo]. [tbl_employees_stage1_" + timeStamp + "]";
            string query1 = "select * into" + " " + destinationTable1 + " " + "from [dbo]. [tbl_employees_stage1]";


            // checking the table is already presnet or not if present returns 1 else return 0
            int connectionresult = 0;
            string checkingtable = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_" + timeStamp + "'";
            SqlDataReader connection = ExecuteQuery(checkingtable, sqlconnection);
            while (connection.Read())
            {
                connectionresult = (int)connection[0];
            }

            //drop backup table if already present
            if (connectionresult == 1)
            {
                SqlCommand cmd = new SqlCommand("drop table " + destinationTable1 , sqlconnection);
                cmd.ExecuteNonQuery();
                Console.WriteLine(destinationTable1 + "dropped succesfully");
            }

            // creates the backuptable 1 if not present else it returns error
            ExecuteQuery(query1, sqlconnection);


            int countMainTable = 0;
            string query4 = "Select count(*) from [dbo]. [tbl_employees_stage1]";
            SqlDataReader counter0 = ExecuteQuery(query4, sqlconnection);
            while (counter0.Read())
            {
                countMainTable = (int)counter0[0];
            }

            int countMainTableBackup = 0;
            string query5 = "Select count(*) from " + destinationTable1;
            SqlDataReader counter1 = ExecuteQuery(query5, sqlconnection);
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
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            #endregion

            #region MyRegion

            string timeStamp2 = DateTime.Now.ToString("MMddyyyy");
            string destinationTable2 = "[dbo]. [tbl_employees_stage1_hold_" + timeStamp2 + "]";
            string backupquery = "select * into" + " " + destinationTable2 + " " + "from [dbo]. [tbl_employees_stage1_hold]";


            // checking the table is already presnet or not if present returns 1 else return 0
            int connectionresult2 = 0;
            string checkingtable2 = "SELECT 1 FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = " + "'tbl_employees_stage1_hold_" + timeStamp2 + "'";
            
            SqlDataReader connection2 = ExecuteQuery(checkingtable2, sqlconnection);
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
            ExecuteQuery(backupquery, sqlconnection);


            int countMainTable2 = 0;
            string backupquery4 = "Select count(*) from tbl_employees_stage1_hold";
            SqlDataReader backupcounter1 = ExecuteQuery(backupquery4, sqlconnection);
            while (counter0.Read())
            {
                countMainTable2 = (int)backupcounter1[0];
            }

            int countMainTableBackup2 = 0;
            string backquery5 = "Select count(*) from " + destinationTable2;
            SqlDataReader backupcounter2 = ExecuteQuery(backquery5, sqlconnection);
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
            Console.WriteLine("--------------------------------------------------------------------------------------------");

            #endregion

            #region MyRegion
            string Query3 = "Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid";
            SqlDataReader datareader = ExecuteQuery(Query3, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
            }
            Console.WriteLine("------------------Query3 is executed----------------");
            #endregion

            #region MyRegion
            string Query4 = "select uniqueid,count(uniqueid),datasourceid from tbl_employees_import group by uniqueid,datasourceid having count(uniqueid)>1";
            SqlDataReader datareader4 = ExecuteQuery(Query4, sqlconnection);
            while (datareader4.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
            }

            Console.WriteLine("------------------Query4 is executed----------------");

            #endregion

            sqlconnection.Close();
            

            //Read Excel  people Report File
            ReadExcelFile();


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

        public static void ReadExcelFile()
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = true;

            string path = "C:\\Users\\kwr579\\Desktop\\AUTOMATION\\People Report_1215.xlsx";
            Workbook wb;
            Worksheet ws;

            try
            {
                wb = app.Workbooks.Open(path);
                ws = wb.Worksheets["sheet1"];

                string cellData = " " + ws.Range["A1"].Value;
                Console.WriteLine(cellData);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
            

            Console.ReadLine();
        }
        //vfhdivhdfvuidnfivn
        //cwdfwevwv
    }
}
