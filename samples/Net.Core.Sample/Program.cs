using ClosedXML.Excel;
using ClosedXML.TableReader.Model;
using Net.Core.Sample.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using Path = System.IO.Path;

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

            IEnumerable<SimpleTable> data = wb.ReadTable<SimpleTable>(1);


            var l = wb.ReadTable<Models.TableWithHeaders>(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = true
                });

            Console.WriteLine(l.ToList().Count);


            //Example with trasnformations
            Expression<Func<string, DateTime>> fConvertData = _ => DateTime.Now.AddDays(10);
            Expression<Func<string, bool>> fSelectToBool = c => c == "x";

            var ls = wb.ReadTable<Models.TableWithHeadersAndBoleanConversion>(1,
                new ReadOptions()
                {
                    TitlesInFirstRow = true,
                    Converters = new Dictionary<string, LambdaExpression>()
                    {
                        {nameof(TableWithHeadersAndBoleanConversion.Birthday),fConvertData },
                        {nameof(TableWithHeadersAndBoleanConversion.Selected),fSelectToBool },
                    }
                });

            wb.Dispose();

            ImportPersonas();


            Console.WriteLine(ls.ToList().Count);

            Console.ReadLine();
        }



        private static void ImportPersonas()
        {
            var s = GetExcelPath("ExcelPrueba.xlsx");
            var bytes = GetExcel(s);
            var wb = new ClosedXML.Excel.XLWorkbook(new MemoryStream(bytes));


            Expression<Func<string, Guid>> fConvertGuid = d => Guid.Parse(d);
            Expression<Func<string, Guid>> fConvertGuidNullable = d => d != null ? Guid.Parse(d) : Guid.Empty;
            Expression<Func<string, bool>> fSelectTBool = c => c == "1";
            Expression<Func<string, int?>> fSexo = sexoString => string.IsNullOrEmpty(sexoString)
                ? (int?)null
                : sexoString.ToLower() == "femenino"
                    ? 1
                    : 0;

            IEnumerable<PersonaImportada> personasImportadas = wb.ReadTable<PersonaImportada>(2
                , new ReadOptions()
                {
                    TitlesInFirstRow = true,
                    Converters = new Dictionary<string, LambdaExpression>()
                    {
                        {nameof(PersonaImportada.IdFuac), fConvertGuid},
                        {nameof(PersonaImportada.TipoDocumentoId), fConvertGuidNullable},
                        {nameof(PersonaImportada.TipoViaId), fConvertGuidNullable},
                        {nameof(PersonaImportada.TelefonoMovilAdmiteSms), fSelectTBool},
                        {nameof(PersonaImportada.RecibirBoletin), fSelectTBool},
                        {nameof(PersonaImportada.AceptoPoliticas), fSelectTBool},
                        {nameof(PersonaImportada.Bloqueado), fSelectTBool},
                        {nameof(PersonaImportada.SexoStr),fSexo}
                    }
                });


            foreach (var personasImportada in personasImportadas)
            {
                ObjectDumper.ObjectDumperExtensions.DumpToString(personasImportada, "persona");
            }

            wb.Dispose();
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
