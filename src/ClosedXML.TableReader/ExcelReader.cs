using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using ClosedXML.Excel;
using ClosedXML.TableReader.Model;

namespace ClosedXML.TableReader;

[Obsolete("Use TableReader extensions methods")]
public class ExcelReader
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
            var ws = WorkBook.Worksheet(sheetNumber);
            return ReadExcelSheet(ws, options);
        }
    }


    public IEnumerable<T> ReadExcelSheet<T>(int sheetNumber, ReadOptions options = null) where T : class, new()
    {
        using (MemoryStream ms = new MemoryStream(_file))
        {
            WorkBook = new XLWorkbook(ms);
            var ws = WorkBook.Worksheet(sheetNumber);
            var dt = ReadExcelSheet(ws, options);

            return dt.AsEnumerableTyped<T>(options);
        }
    }




    #region Private Functions

    private DataTable ReadExcelSheet(IXLWorksheet workSheet, ReadOptions options)
    {
        DataTable dt = new DataTable();

        options = options ?? ReadOptions.DefaultOptions;

        //TODO: Implementar opción con columnas sin títulos

        //primera fila de titulos
        bool firstRow = options.TitlesInFirstRow;
        dt.TableName = workSheet.GetNameForDataTable();

        if (options.TitlesInFirstRow)
        {
            //si no tenemos títulos en la tabla utilizamos los nombres de columna del excel para la definición del DataTable
            foreach (var col in workSheet.ColumnsUsed())
            {
                dt.Columns.Add(col.ColumnLetter());
            }
        }


        foreach (IXLRow row in workSheet.RowsUsed())
        {
            //Usamos la primera fila para crear las columnas con los títulos
            //init with options.TitlesInFirstRow 
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
