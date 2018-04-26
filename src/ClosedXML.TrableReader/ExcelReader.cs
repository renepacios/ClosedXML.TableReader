using System.Collections.Generic;
using System.Data;
using System.IO;
using BalNET.Infraestructure.ExcelReader.Contracts;
using BalNET.Infraestructure.ExcelReader.Extensions;
using ClosedXML.Excel;

namespace BalNET.Infraestructure.ExcelReader
{
    public class ExcelReader : IExcelReader
    {
        private readonly byte[] _file;
        protected XLWorkbook WorkBook;

        public ExcelReader()
        {
        }

        public ExcelReader(byte[] file)
        {
            _file = file;
        }


        protected XLWorkbook LoadWB(byte[] file = null)
        {
            if (file == null) file = _file;

            MemoryStream ms = new MemoryStream(file);
            WorkBook = new XLWorkbook(ms);

            return WorkBook;
        }

        // TODO Agregar métodos genéricos para leer excel

        public DataSet ReadExcel(ReadOptions options = null)
        {
            DataSet ds = new DataSet();
            using (MemoryStream ms = new MemoryStream(_file))
            {
                WorkBook = new XLWorkbook(ms);
                foreach (var workBookWorksheet in WorkBook.Worksheets)
                {
                    var dt = ReadExcelSheet(workBookWorksheet, options);

                    ds.Tables.Add(dt);
                }

                return ds;
            }

        }


        public DataTable ReadExcelSheet(int sheetNumber, ReadOptions options = null)
        {
            using (MemoryStream ms = new MemoryStream(_file))
            {
                WorkBook = new XLWorkbook(ms);
                var ws= WorkBook.Worksheet(sheetNumber);
                return ReadExcelSheet(ws, options);
            }
        }


        public IEnumerable<T> ReadExcelSheet<T>(int sheetNumber, ReadOptions options = null) where T : class, new()
        {
            using (MemoryStream ms = new MemoryStream(_file))
            {
                WorkBook = new XLWorkbook(ms);
                var ws = WorkBook.Worksheet(sheetNumber);
                var dt= ReadExcelSheet(ws, options);

                return dt.AsEnumerableTyped<T>();
            }
        }




        #region Private Functions

        private DataTable ReadExcelSheet(IXLWorksheet workSheet, ReadOptions options)
        {
            DataTable dt = new DataTable();

            options = options ?? ReadOptions.DefaultOptions;

            //primera fila de titulos
            bool firstRow = options.TitlesInFirstRow;
            dt.TableName = workSheet.GetNameForDataTable();

            foreach (IXLRow row in workSheet.RowsUsed())
            {
                //Usamos la primera fila para crear las columnas con los títulos
                if (firstRow)
                {
                    foreach (IXLCell cell in row.CellsUsed())
                    {
                        dt.Columns.Add(cell.GetContentWithOutSpaces());
                    }
                    firstRow = false;
                }
                else
                {
                    dt.Rows.Add();
                    int i = 0;

                    foreach (IXLCell cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
                    {
                        dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                        i++;
                    }
                }
            }

            return dt;
        }
        #endregion

    }
}
