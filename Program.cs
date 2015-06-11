using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.IO;

namespace psoptsys_export
{
    class Program
    {
        static void Main(string[] args)
        {

            if(args.Length == 0){
                Console.WriteLine("");
                Console.WriteLine("[How to use]");
                Console.WriteLine("Syntax: psoptsys [Database Path]");
                Console.WriteLine("");
                Console.WriteLine("[Example]");
                Console.WriteLine("psoptsys C:\\database.mdb");
                return;
            }

            string databasePath = args[0].ToString();
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath + ";User Id=admin;Password=;";
            string sqlSelectQuery = "select [Tch_ID],[Tch_FrontName],[Tch_Name],[Tch_SerName] from [Tch_Hist] order by [Tch_ID] asc";

            if (databasePath == string.Empty)
            {
                return;
            }
            else if (CheckFilePath(databasePath) == false)
            {
                Console.WriteLine("Can not find a database file or Path not correct.");
                return;
            }
            else
            {
                OleDbConnection OBC_Connecttion = new OleDbConnection(connectionString);
                OleDbCommand OBC_Command = new OleDbCommand(sqlSelectQuery, OBC_Connecttion);
                OleDbDataReader OBC_DataReader;
                StringBuilder OBC_StrData = new StringBuilder();

                try
                {
                    OBC_Connecttion.Open();
                    OBC_DataReader = OBC_Command.ExecuteReader();

                    int NumberOfColumn = OBC_DataReader.FieldCount;

                    for (int i = 0; i < NumberOfColumn; i++)
                    {
                        OBC_StrData.Append(OBC_DataReader.GetName(i));
                        if (i + 1 < NumberOfColumn)
                        {
                            OBC_StrData.Append(",");
                        }
                    }
                    OBC_StrData.Append(Environment.NewLine);

                    while (OBC_DataReader.Read())
                    {
                        for (int i = 0; i < NumberOfColumn; i++)
                        {
                            OBC_StrData.Append(OBC_DataReader[i].ToString());
                            if (i + 1 < NumberOfColumn)
                            {
                                OBC_StrData.Append(",");
                            }
                        }
                        OBC_StrData.Append(Environment.NewLine);
                    }

                    OBC_StrData.Replace("(อ)", string.Empty);
                    OBC_StrData.Replace("(ป/บ)", string.Empty);
                    OBC_StrData.Replace("(ป.โท)", string.Empty);
                    OBC_StrData.Replace("(ชูชาติ)", string.Empty);

                    using (StreamWriter OBC_SW = new StreamWriter(File.Open("Output.csv", FileMode.Create), Encoding.UTF8))
                    {

                        OBC_SW.Write(OBC_StrData.ToString());
                    }

                    Console.WriteLine("+ Export Data Successfully.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }

                OBC_Connecttion.Close();
            }     
        }

        static bool CheckFilePath(string FilePath)
        {
            Boolean result = File.Exists(FilePath);
            return result;
        }
    }
}
