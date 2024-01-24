using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using static System.Net.Mime.MediaTypeNames;
using System.Data;
using System.Data.SqlClient;


namespace MEHR_Automation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World");
            //Console.ReadLine();

            string configuration = ConfigurationManager.ConnectionStrings["Dbcon"].ToString();
            SqlConnection sqlconnection = new SqlConnection(configuration);
            sqlconnection.Open();

            Console.WriteLine("Connection is successfull");
            SqlCommand cmd = new SqlCommand("Select count (*), datasource, datasourceid from tbl_Employees_Import group by datasource,datasourceid order by datasourceid  ", sqlconnection);
            SqlDataReader datareader = cmd.ExecuteReader();
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + "|" + datareader[2]);
            }

            sqlconnection.Close();
            Console.ReadLine();





        }
    }
}
