using BenchmarkDotNet.Running;
using ExtractExcelApp.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExtractExcels.FastExcellib;
using ExtractExccl = ExtractExcels.FastExcellib.ExtractExccl;
namespace ExtractExcelApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //var summary = BenchmarkRunner.Run<ExtractExccl>();

                var enviroment = System.Environment.CurrentDirectory;
                string projectDirectory = Directory.GetParent(enviroment).Parent.FullName;
                var inputFile = projectDirectory + "\\Input\\Bayer Accounts.xlsx";

                var sheetName = new string[] { "*","SheetName" };
                var columnIndex = 1;
                ExtractExccl extractexcel = new ExtractExccl();

                extractexcel.ExtractDatafromFastExcel(sheetName, columnIndex, inputFile);


            }
            catch (Exception ex )
            {
                throw ex;
            }
            Console.ReadLine();
        }
    }
}