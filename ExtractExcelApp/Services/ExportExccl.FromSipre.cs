using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace ExtractExcelApp.Services
{
    internal class ExportExcclFromSipre
    {
        internal async void ExportDataAndWriteToText(string inputFile)
        {
            try
            {
                var data = await ExportDataToFromExcel(inputFile);
                List<List<string>> subArray = await GetSubArray(data);
                WriteToTextFile(subArray);
            }
            catch (Exception ex)
            {
                throw ex;
            }         
        }
        internal async Task<DataTable> ExportDataToFromExcel(string inputFile) 
        { 
  
            Spire.Xls.Workbook workbookS = new Spire.Xls.Workbook();
            workbookS.LoadFromFile(inputFile);

            //to get sheet
            Spire.Xls.Worksheet sheetS = workbookS.Worksheets[0];

            DataTable dt = new DataTable();
            for (var row = 1; row <= sheetS.Rows.Count(); row++)
            {

                if (row == 1)
                {
                    dt.Columns.Add(sheetS.Range[row, 1].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 2].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 3].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 4].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 5].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 6].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 7].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 8].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 9].Value.Trim().Replace(" ", string.Empty));
                    dt.Columns.Add(sheetS.Range[row, 10].Value.Trim().Replace(" ", string.Empty));

                }
                else
                {
                    dt.Rows.Add(
                        sheetS.Range[row, 1].Value.Trim(),
                        sheetS.Range[row, 2].Value.Trim(),
                        sheetS.Range[row, 3].Value.Trim(),
                        sheetS.Range[row, 4].Value.Trim(),
                        sheetS.Range[row, 5].Value.Trim(),
                        sheetS.Range[row, 6].Value.Trim(),
                        sheetS.Range[row, 7].Value.Trim(),
                        sheetS.Range[row, 8].Value.Trim(),
                        sheetS.Range[row, 9].Value.Trim(),
                        sheetS.Range[row, 10].Value.Trim()
                        );
                }
            }
            sheetS.Dispose();   
            workbookS.Dispose();
           return  dt;
        }
        internal async Task< List<List<string>>> GetSubArray(DataTable dt)
        {
            var accountNumber = (from DataRow row in dt.Rows
                   select new 
                   {
                       AccountNumber = row["AccountNumber"].ToString(),
                   }).Distinct().ToList();


            List<List<string>> chunkArray = new List<List<string>>();
            var chunkSize = 10;

            for (int i = 0; i < accountNumber.Count(); i += 10)
            {
                var subArray = accountNumber.Select(x=>x.AccountNumber).Skip(i).Take(chunkSize).ToList();
                chunkArray.Add(subArray);
            }

            return chunkArray;
        }
        internal async void  WriteToTextFile(List<List<string>> subArray)
        {
            var enviroment = System.Environment.CurrentDirectory;
            string projectDirectory = Directory.GetParent(enviroment).Parent.FullName;
            var path = projectDirectory+"\\OutPut\\AccountNumber.txt";

            File.WriteAllText(path, String.Empty);

            StringBuilder text = new StringBuilder();

            foreach (var itemList in subArray)
            {
                text.AppendLine("\"" + string.Join("\", \"", itemList) + "\""+",");    
                Console.WriteLine(path);
            }
            File.WriteAllText(path, text.ToString());

        }
    }
}
