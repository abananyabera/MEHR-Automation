using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace MEHR_Automation
{
    public class MacroQueries
    {
        ExecuteQueries executeQueries = new ExecuteQueries();
        public void Add_Counts_by_Datasource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 1.Add_Counts_by_Datasource MacroQueries is started for Execution.\n");
            string Query = "With Stage1_count_bydatasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Stage1 INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Stage1.datasourceid = dbo.tbl_Datasources.datasourceid\r\nWHERE (((dbo.tbl_Employees_Stage1.deleted)=0) AND ((dbo.tbl_Employees_Stage1.removed)=0))\r\nGROUP BY dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\n),ImportAdd_count_by_datasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Import_Add.uniquestring) AS CountOfuniquestring, dbo.tbl_Employees_Import_Add.datasourceid, dbo.tbl_Employees_Import_Add.datasource\r\nFROM dbo.tbl_Employees_Import_Add\r\nGROUP BY dbo.tbl_Employees_Import_Add.datasourceid, dbo.tbl_Employees_Import_Add.datasource\r\n)\r\nSELECT Stage1_count_bydatasource.CountOfmasterid AS [Total active], IsNull([CountOfuniquestring],0) AS [Total imported], Stage1_count_bydatasource.datasource, Round((IsNull([CountOfuniquestring],0)/[Countofmasterid])*100,2) AS [Percent added]\r\nFROM Stage1_count_bydatasource LEFT JOIN ImportAdd_count_by_datasource ON Stage1_count_bydatasource.datasourceid = ImportAdd_count_by_datasource.datasourceid\r\n--ORDER BY Round((IsNull([CountOfuniquestring],0)/[Countofmasterid])*100,2) DESC\r\nUNION\r\nSELECT IsNull([CountOfmasterid],0) AS [Total active], ImportAdd_count_by_datasource.CountOfuniquestring AS [Total imported], ImportAdd_count_by_datasource.datasource,  -100 AS [Percent added]\r\nFROM Stage1_count_bydatasource RIGHT JOIN ImportAdd_count_by_datasource ON Stage1_count_bydatasource.datasourceid = ImportAdd_count_by_datasource.datasourceid\r\nWHERE (((Stage1_count_bydatasource.datasourceid) Is Null))\r\nORDER BY [Percent added] DESC;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3]);
            }
            Console.WriteLine("\n Add_Counts_by_Datasource MacroQueries is compl .\n");
        }

        public void Changed_Fields_by_DataSource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 2.Changed_Fields_by_DataSource MacroQuery is started for Execution.\n");
            string Query = "SELECT Count(dbo.tbl_Employees_Import_Changed.masterid) AS CountOfmasterid, \r\ndbo.tbl_Employees_Import_Changed.fieldname, dbo.tbl_Datasources.datasource, dbo.tbl_Employees_Import_Changed.datasourceid\r\nFROM dbo.tbl_Employees_Import_Changed \r\nINNER JOIN dbo.tbl_Datasources \r\nON dbo.tbl_Employees_Import_Changed.datasourceid = dbo.tbl_Datasources.datasourceid\r\nWHERE (((dbo.tbl_Employees_Import_Changed.fieldname) Not In ('orig_state','orig_platform','orig_country','orig_region' ,'orig_division','orig_expatriate')))\r\nGROUP BY dbo.tbl_Employees_Import_Changed.fieldname, dbo.tbl_Datasources.datasource, dbo.tbl_Employees_Import_Changed.datasourceid\r\nORDER BY Count(dbo.tbl_Employees_Import_Changed.masterid) DESC;\r\n";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3]);
            }
            Console.WriteLine("\n Changed_Fields_by_DataSource MacroQueries is completed .\n");
        }

        public void Change_NotUpdated_by_DataSource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 3.Change_NotUpdated_by_DataSource MacroQuery is started for Execution.\n");
            string Query = "SELECT Count(dbo.tbl_Employees_Import_Changed_Not_Updated.masterid) AS CountOfmasterid, \r\ndbo.tbl_Employees_Import_Changed_Not_Updated.fieldname, dbo.tbl_Datasources.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Import_Changed_Not_Updated  \r\nINNER JOIN dbo.tbl_Datasources \r\nON dbo.tbl_Employees_Import_Changed_Not_Updated.datasourceid = dbo.tbl_Datasources.datasourceid\r\nGROUP BY dbo.tbl_Employees_Import_Changed_Not_Updated.fieldname, dbo.tbl_Datasources.datasourceid,\r\ndbo.tbl_Datasources.datasource\r\nORDER BY Count(dbo.tbl_Employees_Import_Changed_Not_Updated.masterid) DESC;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3]);
            }
            Console.WriteLine("\n Change_NotUpdated_by_DataSourceMacroQueries is completed .\n");
        }

        public void AddDeleted_by_DataSource(SqlConnection sqlconnection)
        { 
            Console.WriteLine("\n 4.AddDeleted_by_DataSource MacroQuery is started for Execution.\n");
            string Query = "SELECT Count(dbo.tbl_Employees_Import_Add_Deleted.uniquestring) AS [Total added but formerly deleted], dbo.tbl_Employees_Import_Add_Deleted.datasource\r\nFROM dbo.tbl_Employees_Import_Add_Deleted\r\nGROUP BY dbo.tbl_Employees_Import_Add_Deleted.datasource\r\nORDER BY Count(dbo.tbl_Employees_Import_Add_Deleted.uniquestring) DESC;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1]);
            }
            Console.WriteLine("\nAddDeleted_by_DataSource MacroQueries is completed.\n");
        }

        public void Removed_Countby_DataSource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 5.Removed_Countby_DataSource MacroQuery is started for Execution.\n");
            string Query = "with dsCount_Remove\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Import_Remove.masterid) AS CountOfmasterid, dbo.tbl_Employees_Import_Remove.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Import_Remove INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Import_Remove.datasourceid = dbo.tbl_Datasources.datasourceid\r\nGROUP BY dbo.tbl_Employees_Import_Remove.datasourceid, dbo.tbl_Datasources.datasource\r\n--ORDER BY Count(dbo.tbl_Employees_Import_Remove.masterid) DESC\r\n), Stage1_count_bydatasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Stage1 INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Stage1.datasourceid = dbo.tbl_Datasources.datasourceid\r\nWHERE dbo.tbl_Employees_Stage1.deleted=0 AND dbo.tbl_Employees_Stage1.removed=0\r\nGROUP BY dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\n)\r\nSELECT Removecount.CountOfmasterid AS Removed, Stage1count.CountOfmasterid AS [Total active], Stage1count.datasourceid, Removecount.datasource, Round((removecount.Countofmasterid/stage1count.countofmasterid)*100,2) AS [Percent reduction]\r\nFROM dscount_Remove AS Removecount INNER JOIN Stage1_count_bydatasource AS Stage1count ON Removecount.datasourceid = Stage1count.datasourceid\r\nORDER BY Round((removecount.Countofmasterid/stage1count.countofmasterid)*100,2) DESC;\r\n\r\n";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4]);
            }
            Console.WriteLine("\nRemoved_Countby_DataSource MacroQueries is completed.\n");
        }

        public void Check_Email_Types(SqlConnection sqlconnection) 
        {
            Console.WriteLine("\n 6.Check_Email_Types MacroQuery is started for Execution.\n");
            string Query = "SELECT dbo.tbl_employees_stage1.epassid,dbo.tbl_employees_stage1.First, dbo.tbl_employees_stage1.middle,\r\n  \tdbo.tbl_employees_stage1.Last, dbo.tbl_employees_stage1.internet_email, \r\n  \tdbo.tbl_employees_stage1.email_type_id,  \r\n  \tdbo.tbl_employees_stage1.datasourceid\r\n     \tFROM dbo.tbl_employees_stage1\r\n     \tWHERE (((dbo.tbl_employees_stage1.active)=0) AND \r\n    \t(CharIndex([last],[internet_email],1)=0 )\r\nAND (CharIndex([first],[internet_email],1)=0) AND \r\n     \t((dbo.tbl_employees_stage1.removed)=0)\r\nAND ((dbo.tbl_employees_stage1.deleted)=0))\r\nOR (((dbo.tbl_employees_stage1.internet_email)<>'') AND ((dbo.tbl_employees_stage1.email_type_id)<>1))\r\nand dbo.tbl_employees_stage1.datasourceid not in (7,9)\r\nORDER BY dbo.tbl_employees_stage1.internet_email;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6]);
            }
            Console.WriteLine("\nCheck_Email_Types MacroQueries is completed.\n");
        }

        public void Missing_Email_Types(SqlConnection sqlconnection) 
        {
            Console.WriteLine("\n 7.Missing_Email_Types MacroQuery is started for Execution.\n");
            string Query = "SELECT dbo.tbl_Employees_Stage1.First, dbo.tbl_Employees_Stage1.middle, dbo.tbl_Employees_Stage1.Last, dbo.tbl_Employees_Stage1.country, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Employees_Stage1.internet_email, dbo.tbl_Employees_Stage1.email_type_id, dbo.tbl_Employees_Stage1.active\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.internet_email) Is Not Null And (dbo.tbl_Employees_Stage1.internet_email)<>'') AND ((dbo.tbl_Employees_Stage1.email_type_id) Is Null) AND ((dbo.tbl_Employees_Stage1.active)<>0));";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6] + "|" + datareader[7]);
            }
            Console.WriteLine("\nMissing_Email_Types MacroQueries is completed.\n");
        }

        public void Coastal_Manager_Duplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 8.Coastal_Manager_Duplicates MacroQuery is started for Execution.\n");
            string Query = "SELECT Coastal.uniqueid, Coastal.First, Coastal.Last, Coastal.sbu, [Non Coastal].First,[Non Coastal].Last, [Non Coastal].sbu\r\nFROM ((dbo.tbl_Employees_Stage1 AS Coastal \r\nLEFT JOIN dbo.tbl_Employees_Stage1 AS [Non Coastal] \r\nON (Coastal.last = [Non Coastal].last) \r\nAND (Coastal.first = [Non Coastal].first)) \r\n\r\nLEFT JOIN dbo.tbl_Duplicates ON Coastal.masterid = dbo.tbl_Duplicates.PrimaryMasterid) \r\n\r\nLEFT JOIN dbo.tbl_Duplicates AS tbl_Duplicates_1 ON Coastal.masterid = tbl_Duplicates_1.SecondaryMasterID \r\n\r\nWHERE (((Coastal.uniqueid) Like '%M') AND ((Coastal.datasourceid)=44) AND ((dbo.tbl_Duplicates.PrimaryMasterid) Is Null)\r\nAND ((tbl_Duplicates_1.SecondaryMasterID) Is Null) AND (([Non Coastal].datasourceid)<>44));\r\n";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6]);
            }
            Console.WriteLine("\nCoastal_Manager_Duplicates MacroQueries is completed.\n");
        }

        public void Danisco_in_Workday_Duplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 9.Danisco_in_Workday_Duplicates MacroQuery is started for Execution.\n");
            string Query = "SELECT S1IA.masterid, Stage1.masterid, ImportAdd.datasourceid, ImportAdd.orig_corporate_entity, \r\nStage1.orig_corporate_entity, Stage1.datasourceid, Stage1.orig_internet_email, Stage1.orig_first, Stage1.orig_last\r\n\r\nFROM dbo.tbl_Employees_Stage1 AS Stage1, dbo.tbl_Employees_Import_Add AS ImportAdd \r\nINNER JOIN dbo.tbl_Employees_Stage1 AS S1IA ON ImportAdd.uniquestring = S1IA.uniquestring\r\nWHERE (((S1IA.masterid) Not In (SELECT Primarymasterid from dbo.tbl_Duplicates) And \r\n(S1IA.masterid) Not In (Select Secondarymasterid From dbo.tbl_Duplicates)) AND \r\n((ImportAdd.datasourceid)=63) AND ((Stage1.datasourceid)=77) AND \r\n((Stage1.orig_internet_email)=[ImportAdd].[orig_internet_email] And \r\n(Stage1.orig_internet_email) Is Not Null And (Stage1.orig_internet_email)<>'')\r\n AND ((ImportAdd.orig_country) In ('China','Australia','Singapore','New Zealand','India')))\r\nOR \r\n(((S1IA.masterid) Not In (SELECT Primarymasterid from dbo.tbl_Duplicates) \r\n And (S1IA.masterid) Not In (Select Secondarymasterid From dbo.tbl_Duplicates))\r\n AND ((ImportAdd.datasourceid)=77) AND ((Stage1.datasourceid)=63) \r\nAND ((Stage1.orig_internet_email)=[ImportAdd].[orig_internet_email] \r\nAnd (Stage1.orig_internet_email) Is Not Null And (Stage1.orig_internet_email)<>'')\r\nAND ((ImportAdd.orig_country) In ('China','Australia','Singapore','New Zealand','India'))) \r\n\r\nOR (((S1IA.masterid) Not In (SELECT Primarymasterid from dbo.tbl_Duplicates)\r\nAnd (S1IA.masterid) Not In (Select Secondarymasterid From dbo.tbl_Duplicates))\r\nAND ((ImportAdd.datasourceid)=63) AND ((Stage1.datasourceid)=77) \r\nAND ((Stage1.orig_first)=[ImportAdd].[orig_first]) AND ((Stage1.orig_last)=[ImportAdd].[orig_last]) \r\nAND ((ImportAdd.orig_country) In ('China','Australia','Singapore','New Zealand','India'))) \r\n\r\nOR (((S1IA.masterid) Not In (SELECT Primarymasterid from dbo.tbl_Duplicates)\r\nAnd (S1IA.masterid) Not In (Select Secondarymasterid From dbo.tbl_Duplicates))\r\nAND ((ImportAdd.datasourceid)=77) AND ((Stage1.datasourceid)=63)\r\nAND ((Stage1.orig_first)=[ImportAdd].[orig_first]) AND ((Stage1.orig_last)=[ImportAdd].[orig_last])\r\nAND ((ImportAdd.orig_country) In ('China','Australia','Singapore','New Zealand','India')));\r\n";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6] + "|" + datareader[7] + "|" + datareader[8]);
            }
            Console.WriteLine("\nDanisco_in_Workday_Duplicates MacroQueries is completed.\n");
        }

        public void Pioneer_D_Group_Match_With_DuPont(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 10.Pioneer_D_Group_Match_With_DuPont MacroQuery is started for Execution.\n");
            string Query = "SELECT PIO.masterid, NonPIO.masterid, NonPIO.sbu, PIO.datasourceid, PIO.group_code, PIO.subgroup_Code, PIO.active, NonPIO.active, dbo.tbl_Duplicates.PrimaryMasterid, NonPIO.datasourceid, NonPIO.First, NonPIO.Last,PIO.First, PIO.Last, PIO.sbu, PIO.deleted\r\nFROM (dbo.tbl_Employees_Stage1 AS NonPIO\r\n INNER JOIN (dbo.tbl_Duplicates RIGHT JOIN dbo.tbl_Employees_Stage1 AS PIO ON dbo.tbl_Duplicates.PrimaryMasterid = PIO.masterid)\r\n  ON (NonPIO.last = PIO.last) AND (NonPIO.first = PIO.first)) \r\nLEFT JOIN dbo.tbl_Duplicates AS tbl_Duplicates_1 ON PIO.masterid = tbl_Duplicates_1.SecondaryMasterID\r\nWHERE (((PIO.datasourceid)=27) AND ((PIO.group_code)='D') AND ((PIO.subgroup_Code)<>'NH')\r\n AND ((dbo.tbl_Duplicates.PrimaryMasterid) Is Null) AND ((NonPIO.datasourceid)<>27)\r\n  AND ((PIO.deleted)=0) AND ((tbl_Duplicates_1.SecondaryMasterID) Is Null));";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6] + "|" + datareader[7] + "|" + datareader[8] + "|" + datareader[9] + "|" + datareader[10] + datareader[11] + "|" + datareader[12] + "|" + datareader[13] + "|" + datareader[14] + "|" + datareader[15]);
            }
            Console.WriteLine("\nPioneer_D_Group_Match_With_DuPont MacroQueries is completed.\n");
        }

        public void Potential_Duplicates_with_potential_match(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 11. Potential_Duplicates_with_potential_match MacroQuery is started for Execution.\n");
            string Query = "SELECT Non_Pioneer.masterid, Pioneer.masterid, Non_Pioneer.masterid, Non_Pioneer.First, Non_Pioneer.middle, Non_Pioneer.Last, Pioneer.First, Pioneer.middle, Pioneer.Last, Pioneer.datasourceid, Non_Pioneer.datasourceid, Non_Pioneer.Sbuid, Non_Pioneer.sbu, Pioneer.sbu\r\nFROM dbo.tbl_Employees_Stage1 AS Pioneer INNER JOIN dbo.tbl_Employees_Stage1 AS Non_Pioneer ON (Pioneer.last = Non_Pioneer.last) AND (Pioneer.middle = Non_Pioneer.middle) AND (Pioneer.first = Non_Pioneer.first)\r\nWHERE (((Pioneer.datasourceid)=27) AND ((Non_Pioneer.datasourceid)<>27) AND ((Non_Pioneer.Sbuid)=128 Or (Non_Pioneer.Sbuid)=15 Or (Non_Pioneer.Sbuid)=19) AND ((Pioneer.active)<>0) AND ((Non_Pioneer.active)<>0)) OR (((Pioneer.datasourceid)=27) AND ((Non_Pioneer.datasourceid)<>27) AND ((Non_Pioneer.sbu) Like '%PIO%') AND ((Pioneer.active)<>0) AND ((Non_Pioneer.active)<>0));";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6] + "|" + datareader[7] + "|" + datareader[8] + "|" + datareader[9] + "|" + datareader[10] + datareader[11] + "|" + datareader[12] + "|" + datareader[13]);
            }
            Console.WriteLine("\n Potential_Duplicates_with_potential_match MacroQueries is completed.\n");
        }

        public void MyAccessID_Duplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 12. MyAccessID_Duplicates MacroQuery is started for Execution.\n");
            string Query = "SELECT tbl_Employees_Stage1.masterid, tbl_Employees_Stage1.uniqueid,\r\ntbl_Employees_Stage1.myaccessid, tbl_Employees_Stage1.internet_email,\r\ntbl_Employees_Stage1.email_type_id, tbl_Employees_Stage1.First, tbl_Employees_Stage1.middle,\r\ntbl_Employees_Stage1.Last, tbl_Employees_Stage1.datasourceid, tbl_Employees_Stage1.active,\r\ntbl_Employees_Stage1.removed, tbl_Employees_Stage1.remove_reason, tbl_Employees_Stage1.expat,\r\ntbl_Employees_Stage1.deleted, tbl_Employees_Stage1.sbu, tbl_Employees_Stage1.site,\r\ntbl_Employees_Stage1.importdate, tbl_Employees_Stage1.lastupdated\r\nFROM tbl_Employees_Stage1\r\nWHERE (((tbl_Employees_Stage1.epassid) In\r\n(SELECT [epassid] FROM tbl_Employees_Stage1 As Tmp WHERE \r\n     (active = 1 or (active = 0 and deleted = 0 and removed = 0)) \r\nand epassid is not null and epassid <> '' GROUP BY [epassid] HAVING Count(*)>1 ))\r\nAND ((tbl_Employees_Stage1.removed)<>1) AND((tbl_Employees_Stage1.deleted)<>1))\r\nORDER BY tbl_Employees_Stage1.epassid;";
            SqlDataReader datareader = executeQueries.ExecuteQuery(Query, sqlconnection);
            while (datareader.Read())
            {
                Console.WriteLine(datareader[0] + "|" + datareader[1] + datareader[2] + "|" + datareader[3] + "|" + datareader[4] + "|" + datareader[5] + "|" + datareader[6] + "|" + datareader[7] + "|" + datareader[8] + "|" + datareader[9] + "|" + datareader[10] + datareader[11] + "|" + datareader[12] + "|" + datareader[13] + "|" + datareader[14] + "|" + datareader[15]);
            }
            Console.WriteLine("\n MyAccessID_Duplicates MacroQueries is completed.\n");
        }

        public void Removal_Not_In_Duplicate_Tables(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 13. Removal_Not_In_Duplicate_Tables MacroQuery is started for Execution.\n");
            string Query9 = "SELECT dbo.tbl_Employees_Stage1.*\r\nFROM (dbo.tbl_Employees_Stage1 LEFT JOIN dbo.tbl_Duplicates ON \r\ndbo.tbl_Employees_Stage1.masterid = dbo.tbl_Duplicates.PrimaryMasterid)\r\nLEFT JOIN dbo.tbl_Duplicates AS tbl_Duplicates1 \r\nON dbo.tbl_Employees_Stage1.masterid = tbl_Duplicates1.SecondaryMasterID\r\nWHERE (tbl_Employees_Stage1.removed=1  \r\nAND ((dbo.tbl_Employees_Stage1.deleted)=0) AND \r\n((dbo.tbl_Duplicates.PrimaryMasterid) Is Null) AND \r\n((tbl_Duplicates1.SecondaryMasterID) Is Null)\r\n AND ((dbo.tbl_Employees_Stage1.datasourceid)<>16));";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n Removal_Not_In_Duplicate_Tables MacroQueries is completed.\n");
        }

        public void Email_Duplicates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 14. Email_Duplicates  MacroQuery is started for Execution.\n");
            string Query9 = "Select dbo.tbl_Employees_Stage1.uniqueid, dbo.tbl_Employees_Stage1.epassid,\r\ndbo.tbl_Employees_Stage1.orig_epassid, dbo.tbl_Employees_Stage1.epassstatusid, dbo.tbl_Employees_Stage1.internet_email, \r\ndbo.tbl_Employees_Stage1.email_type_id, dbo.tbl_Employees_Stage1.First, dbo.tbl_Employees_Stage1.middle, \r\ndbo.tbl_Employees_Stage1.Last, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Employees_Stage1.active, \r\ndbo.tbl_Employees_Stage1.removed, dbo.tbl_Employees_Stage1.remove_reason, dbo.tbl_Employees_Stage1.expat, \r\ndbo.tbl_Employees_Stage1.deleted, dbo.tbl_Employees_Stage1.sbu, dbo.tbl_Employees_Stage1.site, \r\ndbo.tbl_Employees_Stage1.importdate, dbo.tbl_Employees_Stage1.lastupdated, dbo.tbl_Employees_Stage1.myaccessid\r\nFROM dbo.tbl_Employees_Stage1\r\nWHERE (((dbo.tbl_Employees_Stage1.internet_email) In (SELECT internet_email FROM dbo.tbl_Employees_Stage1 As Tmp \r\nWHERE (active = 1 or (active = 0 and deleted = 0 and removed = 0)) and internet_email is not null and internet_email <> '' \r\nand (email_type_id is null or email_type_id = 1) GROUP BY internet_email HAVING Count(*)>1 )) AND \r\n((dbo.tbl_Employees_Stage1.removed)<>1) AND ((dbo.tbl_Employees_Stage1.deleted)<>1))\r\nORDER BY dbo.tbl_Employees_Stage1.internet_email;";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n Email_Duplicates MacroQueries is completed.\n");
        }

        public void Add_Delete_Expatriates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 15. Add_Delete_Expatriates  MacroQuery is started for Execution.\n");
            string Query9 = "SELECT dbo.tbl_Employees_Import_Add_Deleted.*\r\nFROM dbo.tbl_Employees_Import_Add_Deleted\r\nWHERE (((dbo.tbl_Employees_Import_Add_Deleted.orig_expatriate)<>'NO' \r\nAnd Not (dbo.tbl_Employees_Import_Add_Deleted.orig_expatriate) Is Null \r\nAnd (dbo.tbl_Employees_Import_Add_Deleted.orig_expatriate)<>'')) \r\nOR (((dbo.tbl_Employees_Import_Add_Deleted.orig_employment_type) Like '%expat%' Or (dbo.tbl_Employees_Import_Add_Deleted.orig_employment_type) Like '%overseas%') AND ((dbo.tbl_Employees_Import_Add_Deleted.datasourceid)=7));";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n Add_Delete_Expatriates MacroQueries is completed.\n");
        }

        public void New_Expatriates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 16. New_Expatriates  MacroQuery is started for Execution.\n");
            string Query9 = "SELECT dbo.tbl_Employees_Import_Add.*, dbo.tbl_Employees_Import_Add.orig_employment_type, \r\ndbo.tbl_Employees_Import_Add.datasourceid\r\nFROM dbo.tbl_Employees_Import_Add INNER JOIN dbo.tbl_Employees_Stage1 ON dbo.tbl_Employees_Import_Add.uniquestring = dbo.tbl_Employees_Stage1.uniquestring\r\nWHERE ((Not (dbo.tbl_Employees_Import_Add.orig_expatriate) Is Null And \r\n(dbo.tbl_Employees_Import_Add.orig_expatriate)<>'' And (dbo.tbl_Employees_Import_Add.orig_expatriate)<>'No' \r\nAnd (dbo.tbl_Employees_Import_Add.orig_expatriate)<>'Not applicable' And (dbo.tbl_Employees_Import_Add.orig_expatriate)<>'N/A')) OR \r\n(((dbo.tbl_Employees_Import_Add.orig_employment_type) Like '%Expat%' Or \r\n(dbo.tbl_Employees_Import_Add.orig_employment_type) Like '%Overseas%') AND ((dbo.tbl_Employees_Import_Add.datasourceid)=7));";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n New_Expatriates MacroQueries is completed.\n");
        }

        public void Removed_Expatriates(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 17. Removed_Expatriates  MacroQuery is started for Execution.\n");
            string Query9 = "SELECT dbo.tbl_Employees_Stage1.*\r\nFROM dbo.tbl_Employees_Import_Remove INNER JOIN dbo.tbl_Employees_Stage1 ON dbo.tbl_Employees_Import_Remove.masterid = dbo.tbl_Employees_Stage1.masterid\r\nWHERE (((dbo.tbl_Employees_Stage1.deleted)= 1) AND ((dbo.tbl_Employees_Import_Remove.expat)=1));";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n Removed_Expatriates MacroQueries is completed.\n");
        }

        public void vw_AddCountByDataSource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 18. vw_AddCountByDataSource  MacroQuery is started for Execution.\n");
            string Query9 = "With Stage1_count_bydatasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Stage1 INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Stage1.datasourceid = dbo.tbl_Datasources.datasourceid\r\nWHERE (((dbo.tbl_Employees_Stage1.deleted)=0) AND ((dbo.tbl_Employees_Stage1.removed)=0))\r\nGROUP BY dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\n),ImportAdd_count_by_datasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Import_Add.uniquestring) AS CountOfuniquestring, dbo.tbl_Employees_Import_Add.datasourceid, dbo.tbl_Employees_Import_Add.datasource\r\nFROM dbo.tbl_Employees_Import_Add\r\nGROUP BY dbo.tbl_Employees_Import_Add.datasourceid, dbo.tbl_Employees_Import_Add.datasource\r\n)\r\nSELECT Stage1_count_bydatasource.CountOfmasterid AS [Total active], IsNull([CountOfuniquestring],0) AS [Total imported], Stage1_count_bydatasource.datasource, Round((IsNull([CountOfuniquestring],0)/[Countofmasterid])*100,2) AS [Percent added]\r\nFROM Stage1_count_bydatasource LEFT JOIN ImportAdd_count_by_datasource ON Stage1_count_bydatasource.datasourceid = ImportAdd_count_by_datasource.datasourceid\r\n--ORDER BY Round((IsNull([CountOfuniquestring],0)/[Countofmasterid])*100,2) DESC\r\nUNION\r\nSELECT IsNull([CountOfmasterid],0) AS [Total active], ImportAdd_count_by_datasource.CountOfuniquestring AS [Total imported], ImportAdd_count_by_datasource.datasource,  -100 AS [Percent added]\r\nFROM Stage1_count_bydatasource RIGHT JOIN ImportAdd_count_by_datasource ON Stage1_count_bydatasource.datasourceid = ImportAdd_count_by_datasource.datasourceid\r\nWHERE (((Stage1_count_bydatasource.datasourceid) Is Null))\r\nORDER BY [Percent added] DESC;";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n vw_AddCountByDataSource MacroQueries is completed.\n");
        }

        public void vw_RemoveCountByDatasource(SqlConnection sqlconnection)
        {
            Console.WriteLine("\n 19. vw_RemoveCountByDatasource  MacroQuery is started for Execution.\n");
            string Query9 = "with dsCount_Remove\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Import_Remove.masterid) AS CountOfmasterid, dbo.tbl_Employees_Import_Remove.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Import_Remove INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Import_Remove.datasourceid = dbo.tbl_Datasources.datasourceid\r\nGROUP BY dbo.tbl_Employees_Import_Remove.datasourceid, dbo.tbl_Datasources.datasource\r\n--ORDER BY Count(dbo.tbl_Employees_Import_Remove.masterid) DESC\r\n), Stage1_count_bydatasource\r\nas\r\n(\r\nSELECT Count(dbo.tbl_Employees_Stage1.masterid) AS CountOfmasterid, dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\nFROM dbo.tbl_Employees_Stage1 INNER JOIN dbo.tbl_Datasources ON dbo.tbl_Employees_Stage1.datasourceid = dbo.tbl_Datasources.datasourceid\r\nWHERE dbo.tbl_Employees_Stage1.deleted=0 AND dbo.tbl_Employees_Stage1.removed=0\r\nGROUP BY dbo.tbl_Employees_Stage1.datasourceid, dbo.tbl_Datasources.datasource\r\n)\r\nSELECT Removecount.CountOfmasterid AS Removed, Stage1count.CountOfmasterid AS [Total active], Stage1count.datasourceid, Removecount.datasource, Round((removecount.Countofmasterid/stage1count.countofmasterid)*100,2) AS [Percent reduction]\r\nFROM dscount_Remove AS Removecount INNER JOIN Stage1_count_bydatasource AS Stage1count ON Removecount.datasourceid = Stage1count.datasourceid\r\nORDER BY Round((removecount.Countofmasterid/stage1count.countofmasterid)*100,2) DESC;";
            SqlDataReader datareader9 = executeQueries.ExecuteQuery(Query9, sqlconnection);
            while (datareader9.Read())
            {
                Console.WriteLine(datareader9[0] + "|" + datareader9[1]);
            }
            Console.WriteLine("\n vw_RemoveCountByDatasource MacroQueries is completed.\n");
        }

    }
}
