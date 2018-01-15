using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Grids2Csv
{
    class Program
    {
        static void Main(string[] args)
        {
            string line = string.Empty;
            do
            {
                Console.WriteLine("Input Grid xlsx.");
                line = Console.ReadLine();
                if (!File.Exists(line))
                {
                    Console.WriteLine("File not exsits.");
                    continue;
                }

                try
                {
                    WriteCSV(line);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message + ex.StackTrace);
                }
                
            } while (line.ToLower() != "exit");
        }

        private static void WriteCSV(string line)
        {
            var dir = Path.GetDirectoryName(line);
            var dataTable = ReadExcelToTable(line);

            var headerRow = dataTable.Rows[0];
            int row = int.Parse(headerRow[1].ToString());
            int col = int.Parse(headerRow[3].ToString());
            int zIndex = int.Parse(headerRow[5].ToString());
            int resolution = int.Parse(headerRow[7].ToString());
            double heightR = double.Parse(headerRow[9].ToString());
            int span = row*col;

            var outPath = GetOutPath(dir);
            using (var fs = new FileStream(outPath, FileMode.Create))
            {
                using (var sw = new StreamWriter(fs))
                {
                    sw.WriteLine("X,Y,Z,Index");

                    int index = 0;
                    for (int i = 1; i < 1 + row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            List<string> results = new List<string>();
                            //X
                            results.Add((j*resolution).ToString());
                            //Y
                            results.Add(((row - i + 1)*resolution).ToString());
                            //Z
                            SetHeight(zIndex, results, heightR);
                            //Index
                            results.Add((index + zIndex*span).ToString());

                            //Value
                            string value = dataTable.Rows[i][j].ToString();
                            if (!string.IsNullOrEmpty(value))
                                results.AddRange(value.Split(';'));

                            sw.WriteLine(String.Join(",", results));
                            index++;
                        }
                    }
                }
            }
        }

        private static void SetHeight(int zIndex, List<string> results, double heightR)
        {
            if (zIndex == 1)
            {
                results.Add((heightR/2).ToString("F2"));
            }
            else if (zIndex == 0)
            {
                results.Add(0.ToString());
            }
            else
            {
                results.Add(((heightR/2) + (zIndex - 1)*heightR).ToString("F2"));
            }
        }

        private static string GetOutPath(string dirPath)
        {
            const string NAME = "Grids";
            int index = 0;
            do
            {
                string outPath = Path.Combine(dirPath, NAME + index + ".csv");
                if (!File.Exists(outPath))
                {
                    return outPath;
                }

                index++;
            } while (true);
        }

        public static DataTable ReadExcelToTable(string path)//excel存放的路径
        {
            try
            {

                //连接字符串
                string connstring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';"; // Office 07及以上版本 不能出现多余的空格 而且分号注意
                using (OleDbConnection conn = new OleDbConnection(connstring))
                {
                    conn.Open();
                    DataTable sheetsName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" }); //得到所有sheet的名字
                    string firstSheetName = sheetsName.Rows[0][2].ToString(); //得到第一个sheet的名字
                    string sql = string.Format("SELECT * FROM [{0}]", firstSheetName); //查询字符串

                    OleDbDataAdapter ada = new OleDbDataAdapter(sql, connstring);
                    DataSet set = new DataSet();
                    ada.Fill(set);
                    return set.Tables[0];
                }
            }
            catch (Exception ex)
            {
                return null;
            }

        }
    }
}
