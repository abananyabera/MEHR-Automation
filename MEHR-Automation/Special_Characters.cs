using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace MEHR_Automation
{
    public class Special_Characters
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        StoredProcedure StoredProcedure = new StoredProcedure();

        public void findSpecialChars(SqlConnection sqlconnection)
        {
            Console.WriteLine("\nstored procedure findSpeciaChar started ");
            List<char> specialCharacters = new List<char>();
            string Query = "exec find_Specialchar";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader.GetName(0), dataReader.GetName(1), dataReader.GetName(2), dataReader.GetName(3));
            if (dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader[0], dataReader[1], dataReader[2], dataReader[3]);
                }
                Console.WriteLine("  ** PLEASE UPDATE IF THERE ARE ANY SPECIAL CHARACTERS THAT NEED TO BE UPDATED IF ANY MANUALLY ** ");
            }
        }

        public void findSpecialChar(SqlConnection sqlconnection)
        {
            Console.WriteLine("\nstored procedure findSpeciaChar started ");
            List<char> specialCharacters = new List<char>();
            string Query = "exec find_Specialchar";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader.GetName(0), dataReader.GetName(1), dataReader.GetName(2), dataReader.GetName(3));
            if (!dataReader.HasRows)
            {
                while (dataReader.Read())
                {
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader[0], dataReader[1], dataReader[2], dataReader[3]);
                }
                Console.WriteLine("\n No special character update is requires since it has no data present ");
            }
            else
            {
                for (int i = 0; i <= 255; i++)
                {
                    char c = (char)i;
                    if (!char.IsLetterOrDigit(c) && !char.IsWhiteSpace(c) && c != ' ' && c != ',' && c != '-')
                    {
                        specialCharacters.Add(c);
                    }
                }
                
                while (dataReader.Read())
                {
                    Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15}", dataReader[0], dataReader[1], dataReader[2], dataReader[3]);
                    char letter = Convert.ToChar(dataReader[2]);
                    int masterid = Convert.ToInt32(dataReader[0]);
                    string Field = Convert.ToString(dataReader[3]);
                    string CountryId = Convert.ToString(dataReader[1]);
                    foreach (char i in specialCharacters)
                    {
                        if (letter == i)
                        {
                            Console.WriteLine( "\n Hello");
                            selectQuery(masterid, letter, Field,CountryId,sqlconnection);
                        }

                    }
                }

            }
            Console.WriteLine("stored procedure findSpecialChar is Executed");
        }

        public void selectQuery(int masterid, char letter, string Field ,string CountryId,SqlConnection sqlconnection)
        {
            string Query = "SELECT * FROM tbl_employees_stage1 WHERE masterid in (" + masterid + ")";
            SqlDataReader selectQuerydatareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15} | {8,-15}", selectQuerydatareader.GetName(0), selectQuerydatareader.GetName(3), selectQuerydatareader.GetName(4), selectQuerydatareader.GetName(5), selectQuerydatareader.GetName(6), selectQuerydatareader.GetName(7), selectQuerydatareader.GetName(8), selectQuerydatareader.GetName(23), selectQuerydatareader.GetName(24), selectQuerydatareader.GetName(28), selectQuerydatareader.GetName(29));
            while (selectQuerydatareader.Read())
            {
                Console.WriteLine("{0,-15} | {1,-15} | {2,-15} | {3,-15} | {4,-15} | {5,-15} | {6,-15} | {7,-15}|{8,-15}",
                    selectQuerydatareader[0], selectQuerydatareader[3], selectQuerydatareader[4],
                    selectQuerydatareader[5], selectQuerydatareader[6], selectQuerydatareader[7],
                    selectQuerydatareader[8], selectQuerydatareader[23], selectQuerydatareader[24]);
                    
                string internet_email = Convert.ToString(selectQuerydatareader[23]);
                check_special_char(masterid, letter, Field, CountryId, internet_email, sqlconnection);
            }
        }

        public void check_special_char(int masterid, char letter, string Field, string CountryId, string internet_email, SqlConnection sqlconnection)
        {
            string Query = "SELECT * FROM map_specchar where spchar = '" + letter + "'";
            SqlDataReader dataReader = executeQueries.ExecuteQuery(Query, sqlconnection);
            if(!dataReader.HasRows)
            {
                Console.WriteLine(" Special character is not present  Insert the special character");
                string insert_Query = "INSERT INTO map_specchar values ('"+ letter +"',"+ CountryId +",'" + internet_email+",'";
                Console.WriteLine(insert_Query);
                StoredProcedure.replaceSpecialChar(sqlconnection);

            }
            else
            {
                Console.WriteLine(" Special character is already present ");
            }
        }
    }
}
