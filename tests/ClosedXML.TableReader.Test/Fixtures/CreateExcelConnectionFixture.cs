using ClosedXML.Excel;
using System;
using System.IO;

namespace ClosedXML.TableReader.Test.Fixtures
{
    public class CreateExcelConnectionFixture : IDisposable
    {
        private XLWorkbook _wb;
        public XLWorkbook Wb
        {
            get => _wb ?? throw new NullReferenceException("You need call Init before Use this fixture");
            private set => _wb = value;
        }

        public XLWorkbook Init(string documentName)
        {

            if (Path.GetExtension(documentName) == string.Empty) documentName += ".xlsx";

            var excelPath = Path.Combine(Directory.GetCurrentDirectory(), "ExcelDocuments", documentName);
            byte[] bytes = System.IO.File.ReadAllBytes(excelPath);

            Wb = new ClosedXML.Excel.XLWorkbook(new MemoryStream(bytes));

            return Wb;
        }




        private void ReleaseUnmanagedResources()
        {
            Wb.Dispose();

        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~CreateExcelConnectionFixture()
        {
            ReleaseUnmanagedResources();
        }
    }
}
