using FastExcel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using Formatting = Newtonsoft.Json.Formatting;

namespace ExtractExcels.FastExcellib
{

    public class ExtractExccl
    {
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
                    var rowRecords = worksheet.Rows?.Skip(columnIndex - 1).ToList();

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
                }
            }
        

        JsonSerializerSettings jsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.Indented
        };

            jsonData = JsonConvert.SerializeObject(worksheetData, jsonSettings);
            Console.WriteLine(jsonData);
            return jsonData;
        }

        public bool HasSheetsName(string[] givenSheetName, string workSheetName)
        {
            if (givenSheetName.Where(x => x.Equals("*", StringComparison.OrdinalIgnoreCase)).Any())
            {
                return true;
            }
            if (givenSheetName.Any(x => x.Equals(workSheetName, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return false;
        }
    }
}
