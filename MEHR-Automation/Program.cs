using System;
using System.Configuration;
using System.Data.SqlClient;
using static System.Net.Mime.MediaTypeNames;

namespace MEHR_Automation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World");

            // Retrieve connection string from app.config
            //string connectionString = "Data Source = localhost\\SQLEXPRESS; Database = AppData; Initial Catalog = test; Integrated Security = SSPI"; 
            string connectionString = ConfigurationManager.ConnectionStrings["Dbcon"].ConnectionString;


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    // Open the database connection
                    connection.Open();
                    // Perform database operations here
                    Console.WriteLine("Connected to the database.");

                    SqlCommand cmd = new SqlCommand("select * from Submission", connection);
                    SqlDataReader datareader = cmd.ExecuteReader();
                    while (datareader.Read())
                    {
                        Console.WriteLine(datareader[0] + " " + datareader["OrganizationName"] + " " + datareader[2]);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message + ex);
                }
            }

            Console.ReadLine();
        }
    }
}
