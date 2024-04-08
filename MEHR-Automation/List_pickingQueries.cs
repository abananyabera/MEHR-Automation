using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MEHR_Automation
{
    public class List_pickingQueries
    {
        PicklingQueries picklingQueries = new PicklingQueries();
        public void List_picking_Queries(SqlConnection sqlconnection)
        {
            Console.WriteLine("Next Action : Please click Enter to execute the 'UnMappedEntities' Pickling Query");
            ReadLine();
            picklingQueries.UnMappedEntities(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to execute the 'UnmappedSBUs' Pickling Query");
            ReadLine();
            picklingQueries.UnmappedSBUs(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to execute the 'UnmappedSites' Pickling Query");
            ReadLine();
            picklingQueries.UnmappedSites(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to execute the 'UnmappedOps' Pickling Query");
            ReadLine();
            picklingQueries.UnmappedOps(sqlconnection);

            Console.WriteLine("Next Action : Please click Enter to execute the 'UnmappedFunctions' Pickling Query");
            ReadLine();
            picklingQueries.UnmappedFunctions(sqlconnection);
        }

        public static void ReadLine()
        {
            Console.WriteLine("-------------------------------------------------------------");
            Console.ReadLine();
        }
    }
}
