using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using ClosedXML.Excel;
using ClosedXML.TableReader;
using ClosedXML.TableReader.Model;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using Net.Core.Sample.Models;

namespace Net.Core.Sample
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World!");

            var s = GetExcelPath();
            var bytes = GetExcel(s);
            var wb = new ClosedXML.Excel.XLWorkbook(new MemoryStream(bytes));


            Expression<Func<string, DateTime>> f = _ => DateTime.Now.AddDays(10);

            var l = wb.ReadTable<Models.TableWithHeaders>(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = true,

                    Converters = new Dictionary<string, LambdaExpression>()
                    {
                        {nameof(TableWithHeaders.Birthday), f}
                    }
                });

            Console.WriteLine(l.ToList().Count);

            Console.ReadLine();
        }



        private static byte[] GetExcel(string excelPath)
        {
            byte[] file = System.IO.File.ReadAllBytes(excelPath);

            return file;
        }

        private static string GetExcelPath(string name = "TableWithHeaders.xlsx")
        {
            return Path.Combine(Directory.GetCurrentDirectory(), "Excels", name);
        }
    }
}
