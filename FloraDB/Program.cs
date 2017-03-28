namespace FloraDB
{
    using System;
    using System.Collections.Generic;
    using System.Configuration;
    using System.IO;
    using System.Net;
    using System.Reflection;
    using System.Runtime.InteropServices;
    using System.Text;
    using System.Threading.Tasks;
    using System.Web;

    using Microsoft.Office.Interop.Excel;

    internal class Program
    {
        private static string GetPlantRecords(string name)
        {
            name = name.Trim();
            Console.WriteLine($"Searching for {name}..");
            var webRequest =
                WebRequest.Create(
                    "http://www.ipni.org/ipni/simplePlantNameSearch.do?find_wholeName=" + HttpUtility.UrlEncode(name)
                    + "&output_format=delimited-short");
            var response = webRequest.GetResponse();
            return Encoding.UTF8.GetString(GetStreamBytes(response.GetResponseStream()));
        }

        private static byte[] GetStreamBytes(Stream stream)
        {
            var array = new byte[16384];
            byte[] result;
            using (var memoryStream = new MemoryStream())
            {
                int count;
                while ((count = stream.Read(array, 0, array.Length)) > 0) memoryStream.Write(array, 0, count);
                result = memoryStream.ToArray();
            }

            return result;
        }

        private static void Main(string[] args)
        {
            var application =
                (Application)
                Activator.CreateInstance(Marshal.GetTypeFromCLSID(new Guid("00024500-0000-0000-C000-000000000046")));
            var workbook = application.Workbooks.Open(
                ConfigurationManager.AppSettings["fileName"],
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                Type.Missing);

            Worksheet worksheet = workbook.Sheets["Sheet1"];
            Worksheet worksheet2 = workbook.Sheets["Sheet2"];
            var usedRange = worksheet.UsedRange;
            var num = 0;
            var num2 = 0;

            List<Task<List<string[]>>> tasks = new List<Task<List<string[]>>>();

            foreach (Range range in usedRange.Rows)
            {
                var row = range.Row;
                object obj = (range.Cells[1, Missing.Value] as Range).Value;
                if (obj != null)

                    tasks.Add(new Task<List<string[]>>(() =>
                                {
                                    var data = GetPlantRecords(obj as string);
                                    return ParseResult(data);
                                }));
            }

            foreach (var task in tasks)
                task.Start();
            Task.WaitAll(tasks.ToArray());
            int counter = 0;

            foreach (Range range in usedRange.Rows)
            {
                var row = range.Row;
                object obj = (range.Cells[1, Missing.Value] as Range).Value;
                if (obj != null)
                {
                    var list = tasks[counter++].Result;
                    try
                    {
                        if(list == null || list.Count ==0)
                            throw new Exception();

                        Console.WriteLine($"Writing {counter}/{tasks.Count}:{obj}..");
                        string text = null;
                        for (var i = num2; i < list.Count; i++)
                            if (num2 == 0)
                            {
                                num++;
                                num2 = 1;
                                for (var j = 0; j < list[i].Length; j++) worksheet2.Cells[num, j + 2] = list[i][j];
                            }
                            else
                            {
                                if (text == null) text = list[i][14].ToLower();
                                else if (list[i][14].ToLower() == text) break;
                                num++;
                                worksheet2.Cells[num, 1] = obj;
                                for (var j = 0; j < list[i].Length; j++) worksheet2.Cells[num, j + 2] = list[i][j];
                            }
                    }
                    catch (Exception)
                    {
                        num++;
                        worksheet2.Cells[num, 1] = obj;
                        worksheet2.Cells[num, 2] = "NOT IN DATABASE!!";
                        Console.WriteLine($"{obj} NOT IN DATABASE!!");
                    }
                }
            }

            workbook.Save();
            workbook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        private static List<string[]> ParseResult(string data)
        {
            List<string[]> result;
            try
            {
                var list = new List<string[]>();
                var array = data.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                for (var i = 0; i < array.Length; i++)
                {
                    var item = array[i].Split('%');
                    list.Add(item);
                }

                result = list;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                result = null;
            }

            return result;
        }
    }
}