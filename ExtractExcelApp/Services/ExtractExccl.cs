using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using FastExcel;
using IronXL;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using OfficeOpenXml;
using Perfolizer.Mathematics.RangeEstimators;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Xml;
using static System.Net.Mime.MediaTypeNames;
using Formatting = Newtonsoft.Json.Formatting;

namespace ExtractExcelApp.Services
{

    public class ExtractExccl
    {
        static object obj = new { };
        internal async void ExportDataAndWriteToText(string inputFile)
        {
            try
            {
                // ExtractDatafromFastExcel();

                //var data = await ExtractDatafromEPPlus(inputFile);
                //List<List<string>> subArray = await GetSubArray(data);
                //WriteToTextFile(subArray);


            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool HasSheetsName(string[] givenSheetName, string workSheetName)
        {
            if (givenSheetName.Where(x => x.Equals("x", StringComparison.OrdinalIgnoreCase)).Any())
            {
                return true;
            }
            if (givenSheetName.Any(x => x.Equals(workSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return false;
        }
        [Benchmark]
        public string ExtractDatafromFastExcel(string[] givenSheetName, int columnIndex, string filePath)
        {

            string jsonData = string.Empty;
            Dictionary<string, DataTable> worksheetData = new Dictionary<string, DataTable>();
            using (var fastExcel = new FastExcel.FastExcel(new FileInfo(filePath), true))  //
            {
                foreach (var worksheet in fastExcel.Worksheets)
                {
                    if (!HasSheetsName(givenSheetName, worksheet.Name))
                    {
                        continue;
                    }

                    worksheet.Read();
                    DataTable dataTable = new DataTable();
                    var columnRecords = worksheet.Rows.Skip(columnIndex - 1).Take(1).FirstOrDefault().Cells.ToList();
                    var rowRecords = worksheet.Rows?.Skip(1).ToList() ?? new List<Row>();

                    if (columnRecords == null)
                    {
                        return string.Empty;
                    }

                    foreach (var item in columnRecords)
                    {
                        dataTable.Columns.Add(item.Value?.ToString() ?? string.Empty);
                    }
                    if (rowRecords != null)
                    {
                        // Add rows to DataTable
                        foreach (var row in rowRecords)
                        {
                            DataRow dataRow = dataTable.Rows.Add();
                            foreach (var cellItem in row.Cells.ToList())
                            {
                                dataRow[cellItem.ColumnNumber - 1] = cellItem.Value ?? string.Empty;
                            }

                        }
                    }
                    if (!worksheetData.ContainsKey(worksheet.Name))
                    {
                        worksheetData.Add(worksheet.Name, dataTable);
                    }


                    // Console.WriteLine(jsonData);

                }
            }
        

        JsonSerializerSettings jsonSettings = new JsonSerializerSettings
        {
            ContractResolver = new CamelCasePropertyNamesContractResolver(),
            Formatting = Formatting.Indented
        };

        jsonData = JsonConvert.SerializeObject(worksheetData, jsonSettings);

            Console.WriteLine(jsonData);
         
            return jsonData;
        }

    //[Benchmark]
    public string ExtractDatafromIronXl()
    {
        string filePath = "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx";
        string jsonData = string.Empty;



        WorkBook workbook = WorkBook.Load(filePath);

        // Assuming the data is in the first worksheet (index 0)
        foreach (var workSheet in workbook.WorkSheets)
        {
            DataTable dataTable = new DataTable();

            var columnRecords = workSheet.Rows.FirstOrDefault()?.ToList();
            var rowRecords = workSheet.Rows.Skip(1).ToList();

            foreach (var item in columnRecords)
            {
                dataTable.Columns.Add(item.Text);

            }
            // Add rows to DataTable
            foreach (var row in rowRecords)
            {
                DataRow dataRow = dataTable.Rows.Add();
                foreach (var cellItem in row.ToList())
                {
                    dataRow[cellItem.ColumnIndex] = cellItem.Value;
                }

            }

            jsonData = DataTableToJson(dataTable);

            //  Console.WriteLine(jsonData);

        }



        return jsonData;
    }

    [Benchmark]
    public async Task<string> ExtractDatafromEPPlus()
    {
        string jsonData = string.Empty;
        string filePath = "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx";

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var package = new ExcelPackage(new System.IO.FileInfo(filePath)))
        {
            // Assuming the data is in the first worksheet (index 1)
            // var worksheet = package.Workbook.Worksheets[0];
            foreach (var worksheet in package.Workbook.Worksheets)
            {

                DataTable dataTable = new DataTable();

                foreach (var firstRowCell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    dataTable.Columns.Add(firstRowCell.Text.Trim().Replace(" ", string.Empty));
                }

                // Start from the second row (index 2) as the first row contains column headers
                for (int rowNum = 2; rowNum <= worksheet.Dimension.End.Row; rowNum++)
                {
                    var row = worksheet.Cells[rowNum, 1, rowNum, worksheet.Dimension.End.Column];
                    DataRow dataRow = dataTable.Rows.Add();
                    foreach (var cell in row)
                    {
                        dataRow[cell.Start.Column - 1] = cell.Text.Trim();
                    }
                }

                jsonData = DataTableToJson(dataTable);
            }


            //Console.WriteLine(jsonData);
        }
        return jsonData;
    }



    [Benchmark]
    public async Task<string> ExportDataFromSpire()
    {
        string jsonData = string.Empty;
        string filePath = "D:\\ExtractExcelApp\\Input\\Bayer Accounts.xlsx";
        //Spire.Xls.Workbook workbookS = new Spire.Xls.Workbook();
        using (Spire.Xls.Workbook workbookS = new Spire.Xls.Workbook())
        {

            workbookS.LoadFromFile(filePath);

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
            jsonData = DataTableToJson(dt);

        }

        return jsonData;
    }

    private string DataTableToJson(DataTable dataTable)
    {
        try
        {
            string json = JsonConvert.SerializeObject(dataTable);

            return json;
        }
        catch (Exception ex)
        {

            throw ex;
        }

        // Convert DataTable to JSON using System.Text.Json

    }

    internal async Task<List<List<string>>> GetSubArray(DataTable dt)
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
            var subArray = accountNumber.Select(x => x.AccountNumber).Skip(i).Take(chunkSize).ToList();
            chunkArray.Add(subArray);
        }

        return chunkArray;
    }
    internal async void WriteToTextFile(List<List<string>> subArray)
    {
        var enviroment = System.Environment.CurrentDirectory;
        string projectDirectory = Directory.GetParent(enviroment).Parent.FullName;
        var path = projectDirectory + "\\OutPut\\AccountNumber.txt";

        File.WriteAllText(path, String.Empty);

        StringBuilder text = new StringBuilder();

        foreach (var itemList in subArray)
        {
            text.AppendLine("\"" + string.Join("\", \"", itemList) + "\"" + ",");
            Console.WriteLine(path);
        }
        File.WriteAllText(path, text.ToString());

    }
}
}
