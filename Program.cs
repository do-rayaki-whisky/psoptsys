using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb; /* ใช้ อ่าน เขียน ฐานข้อมูล */
using System.IO; /* ใช้อ่าน เขียน ข้อมูลลงในไฟล์ */

namespace psoptsys_export
{
    class Program
    {
        static void Main(string[] args)
        {
            /* ตรวจสอบตัวแปร args ว่ามีการใส่ค่าใดๆมาหรือเปล่า ในกรณีที่ไม่ได้มีการกำหนดค่าใดๆมา จะแสดงวิธีการใช้โปรแกรม */
            if(args.Length == 0){
                Console.WriteLine("");
                Console.WriteLine("[How to use]");
                Console.WriteLine("Syntax: psoptsys [Database Path]");
                Console.WriteLine("");
                Console.WriteLine("[Example]");
                Console.WriteLine("psoptsys C:\\database.mdb");
                return;
            }

            string databasePath = args[0].ToString(); /* กำหนดพาธของฐานข้อมูล */
            string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath + ";User Id=admin;Password=;"; /* Connection String Access 2003 */
            string sqlSelectQuery = "select [Tch_ID],[Tch_FrontName],[Tch_Name],[Tch_SerName] from [Tch_Hist] order by [Tch_ID] asc"; /* คำสั้ง SQL เลือกข้อมูล ไอดี คำนำหน้า ชื่อ และนามสกุล */

            if (databasePath == string.Empty)
            {
                return;
            }
            else if (CheckFilePath(databasePath) == false)
            {
                Console.WriteLine("Can not find a database file or Path not correct."); /* ถ้าพาธที่ป้อนเข้ามาไม่ถูกต้องหรือหาไฟล์ไม่เจอ */
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

                    int NumberOfColumn = OBC_DataReader.FieldCount; /* เก็บจำนวนคอลัม */

                    for (int i = 0; i < NumberOfColumn; i++)
                    {
                        OBC_StrData.Append(OBC_DataReader.GetName(i)); /* นำชื่อคอมลัมมาเรียงในรูปแบบ csv ในบรรทัดแรก */
                        if (i + 1 < NumberOfColumn) /* ในคอลัมสุดท้ายจะไม่มีการใส่คอมม่า */
                        {
                            OBC_StrData.Append(",");
                        }
                    }
                    OBC_StrData.Append(Environment.NewLine);

                    while (OBC_DataReader.Read())
                    {
                        for (int i = 0; i < NumberOfColumn; i++)
                        {
                            OBC_StrData.Append(OBC_DataReader[i].ToString()); /* นำข้อมูลทั้งหมดมาใส่ไว้ใน object ในรูปแบบ csv */
                            if (i + 1 < NumberOfColumn)
                            {
                                OBC_StrData.Append(",");
                            }
                        }
                        OBC_StrData.Append(Environment.NewLine);
                    }

                    OBC_StrData.Replace("(อ)", string.Empty); /* ตัดคำที่ไม่ต้องการออก */
                    OBC_StrData.Replace("(ป/บ)", string.Empty); /* ตัดคำที่ไม่ต้องการออก */
                    OBC_StrData.Replace("(ป.โท)", string.Empty); /* ตัดคำที่ไม่ต้องการออก */
                    OBC_StrData.Replace("(ชูชาติ)", string.Empty); /* ตัดคำที่ไม่ต้องการออก */

                    using (StreamWriter OBC_SW = new StreamWriter(File.Open("Output.csv", FileMode.Create), Encoding.UTF8)) /* นำข้อมูลทั้งหมดมาเขียนลงในไฟล์ */
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

        static bool CheckFilePath(string FilePath) /* ฟังชันตรวจสอบพาธของฐานข้อมูล */
        {
            Boolean result = File.Exists(FilePath);
            return result;
        }
    }
}
