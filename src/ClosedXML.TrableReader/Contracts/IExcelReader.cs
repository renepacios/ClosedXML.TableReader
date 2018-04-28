using System.Data;
using ClosedXML.TableReader.Model;

namespace ClosedXML.TableReader.Contracts
{
    public interface IExcelReader
    {
        DataSet ReadExcel(ReadOptions options = null);

        DataTable ReadExcelSheet(int SheetNumber, ReadOptions options = null);
    }
}
