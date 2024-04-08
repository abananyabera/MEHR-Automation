using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class List_MicroQueries
    {
        MacroQueries macroQueries = new MacroQueries();

        public void List_Micro_Queries(SqlConnection sqlconnection)
        {
            //MACROQUERIES EXECUTION
            MacroQueries macroQueries = new MacroQueries();

            Console.WriteLine(" Next Action : Please click Enter to Execute the Add_Counts_by_Datasource Macro Query");
            ReadLine();
            macroQueries.Add_Counts_by_Datasource(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the Changed_Fields_by_DataSource Macro Query");
            ReadLine();
            macroQueries.Changed_Fields_by_DataSource(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the Change_NotUpdated_by_DataSource Macro Query");
            ReadLine();
            macroQueries.Change_NotUpdated_by_DataSource(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the AddDeleted_by_DataSource Macro Query");
            ReadLine();
            macroQueries.AddDeleted_by_DataSource(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the Removed_Countby_DataSource Macro Query");
            ReadLine();
            macroQueries.Removed_Countby_DataSource(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the Check_Email_Types Macro Query");
            ReadLine();
            macroQueries.Check_Email_Types(sqlconnection);

            Console.WriteLine(" Next Action : Please click Enter to Execute the Missing_Email_Types Macro Query");
            ReadLine();
            macroQueries.Missing_Email_Types(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Missing_Email_Types Macro Query");
            ReadLine();
            macroQueries.Coastal_Manager_Duplicates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Danisco_in_Workday_Duplicates Macro Query");
            ReadLine();
            macroQueries.Danisco_in_Workday_Duplicates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Pioneer_D_Group_Match_With_DuPont Macro Query");
            ReadLine();
            macroQueries.Pioneer_D_Group_Match_With_DuPont(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Potential_Duplicates_with_potential_match Macro Query");
            ReadLine();
            macroQueries.Potential_Duplicates_with_potential_match(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the MyAccessID_Duplicates Macro Query");
            ReadLine();
            macroQueries.MyAccessID_Duplicates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Removal_Not_In_Duplicate_Tables Macro Query");
            ReadLine();
            macroQueries.Removal_Not_In_Duplicate_Tables(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Email_Duplicates Macro Query");
            ReadLine();
            macroQueries.Email_Duplicates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Add_Delete_Expatriates Macro Query");
            ReadLine();
            macroQueries.Add_Delete_Expatriates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the New_Expatriates Macro Query");
            ReadLine();
            macroQueries.New_Expatriates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the Removed_Expatriates Macro Query");
            ReadLine();
            macroQueries.Removed_Expatriates(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the vw_AddCountByDataSource Macro Query");
            ReadLine();
            macroQueries.vw_AddCountByDataSource(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to Execute the vw_RemoveCountByDatasource Macro Query");
            ReadLine();
            macroQueries.vw_RemoveCountByDatasource(sqlconnection);
        }

        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }

    }
}
