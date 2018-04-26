using System.Data;
using ClosedXML.Excel;

namespace BalNET.Infraestructure.ExcelReader.Contracts
{
    public interface IExcelReader
    {
        DataSet ReadExcel(ReadOptions options = null);

        DataTable ReadExcelSheet(int SheetNumber, ReadOptions options = null);
    }
}
