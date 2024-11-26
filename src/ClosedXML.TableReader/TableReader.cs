using ClosedXML.TableReader.Attributes;
using ClosedXML.TableReader.Model;
using System.ComponentModel;
using System.Data;
using System.Reflection;

// ReSharper disable once CheckNamespace
namespace ClosedXML.Excel;

public static class TableReader
{
    #region IXLWorkbook Functions

    /// <summary>
    /// Read All Data from Excel into DataSet. Will be create a DataTable by WorkbookSheet
    /// </summary>
    /// <param name="wb">ClosedXml WorkBook instance</param>
    /// <param name="options">Read options <see cref="ReadOptions"/></param>
    /// <returns></returns>
    public static DataSet ReadTables(this IXLWorkbook wb, ReadOptions options = null, CancellationToken cancellationToken = default)
    {
        var ds = new DataSet();

        foreach (var workBookWorksheet in wb.Worksheets)
        {
            var dt = AsDataTable(workBookWorksheet, options, cancellationToken);

            ds.Tables.Add(dt);
        }

        return ds;

    }

    /// <summary>
    /// Read Table from Excel Sheet into a memory DataTable
    /// </summary>
    /// <param name="wb">ClosedXml WorkBook instance</param>
    /// <param name="sheetNumber">Workbook sheet number to read</param>
    /// <param name="options">Read options <see cref="ReadOptions"/></param>
    /// <returns></returns>
    public static DataTable ReadTable(this IXLWorkbook wb, int sheetNumber, ReadOptions options = null, CancellationToken cancellationToken = default)
    {
        if (sheetNumber <= 0 || sheetNumber > wb.Worksheets.Count)
        {
            throw new IndexOutOfRangeException($"{nameof(sheetNumber)} is Out of Range");
        }

        var ws = wb.Worksheet(sheetNumber);
        return ws.AsDataTable(options, cancellationToken);
    }

    /// <summary>
    /// Read Table from Excel Sheet into a typed IEnumerable Collection
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="wb">ClosedXml WorkBook instance</param>
    /// <param name="sheetNumber">Workbook sheet number to read</param>
    /// <param name="options">Read options <see cref="ReadOptions"/></param>
    /// <returns></returns>
    public static IEnumerable<T> ReadTable<T>(this IXLWorkbook wb, int sheetNumber, ReadOptions options = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        if (sheetNumber <= 0 || sheetNumber > wb.Worksheets.Count)
        {
            throw new IndexOutOfRangeException($"{nameof(sheetNumber)} is Out of Range");
        }

        var ws = wb.Worksheet(sheetNumber);
        var dt = ws.AsDataTable(options, cancellationToken);

        return dt.AsEnumerable<T>(options);
    }

    #endregion

    #region IXLWorksheet Functions

    /// <summary>
    /// Read Table from Excel Sheet into a typed IEnumerable Collection
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="workSheet">Workbook sheet number to read</param>
    /// <param name="options">Read options <see cref="ReadOptions"/></param>
    /// <returns></returns>
    public static IEnumerable<T> ReadTable<T>(this IXLWorksheet workSheet, ReadOptions options = null, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        var dt = workSheet.AsDataTable(options, cancellationToken);
        return dt.AsEnumerable<T>(options);
    }

    /// <summary> Read Table from Excel Sheet into a memory DataTable </summary>
    /// <param name="workSheet">Workbook sheet number to read</param>
    /// <param name="options">Read options <see cref="ReadOptions"/></param>
    /// <returns></returns>
    public static DataTable AsDataTable(this IXLWorksheet workSheet, ReadOptions options, CancellationToken cancellationToken)
    {
        var dt = new DataTable();

        options = options ?? ReadOptions.DefaultOptions;

        var initiativeRow = options.RowStart;
        var lastCol = workSheet.LastColumnUsed().ColumnNumber();
        var lastRow = workSheet.LastRowUsed().RowNumber();

        // TODO: verify rows

        //primera fila de títulos
        dt.TableName = workSheet.GetNameForDataTable();

        #region Parse headers

        // foreach (var col in workSheet.ColumnsUsed())
        for (var col = 1; col <= lastCol; col++)
        {
            var cell = workSheet.Cell(initiativeRow, col);
            var colName = cell.Value.ToString();
            var columnLetter = workSheet.ColumnsUsed().ToArray()[col - 1].ColumnLetter();
            AddColumnName(options.TitlesInFirstRow ? colName : columnLetter);

            void AddColumnName(string name, int retryCount = 0)
            {
                try
                {
                    dt.Columns.Add(name);  // add columns to the data table.   
                }
                catch (System.Data.DuplicateNameException)
                {
                    if (retryCount > 100) { throw; }
                    AddColumnName(name + " " + retryCount);
                }
            }
        }

        #endregion

        Func<IXLRow, bool> getRows = (r) => (r.RowNumber() >= (options.TitlesInFirstRow ? (options.RowStart + 1) : options.RowStart));

        foreach (var row in workSheet.RowsUsed(r => getRows(r)))
        {
            cancellationToken.ThrowIfCancellationRequested();

            dt.Rows.Add();
            var i = 0;

            foreach (var cell in row.Cells(row.FirstCellUsed().Address.ColumnNumber, row.LastCellUsed().Address.ColumnNumber))
            {
                dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                i++;
            }
        }

        /*
        var singleDValue = new object[lastCol]; // value array first row contains column names. so loop starts from 2 instead of 1
        for (var row = initiativeRow + 1; row <= lastRow; row++)
        {
            for (var col = initiativeColumn; col < lastCol; col++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    singleDValue[col] = workSheet.Cell(row, col + 1).Value.ToString();
                }
                catch (Exception e)
                {
                    logger.LogError($"Error on cell [row:{row}, col:{col}] {e}");
                    throw new ExcelException(row, col, e.Message, e);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            dt.LoadDataRow(singleDValue, LoadOption.PreserveChanges);
        }
        // */

        return dt;
    }

    #endregion

    #region DataTable Functions

    /// <summary>
    /// Converts a DataTable to a list with generic objects
    /// </summary>
    /// <typeparam name="T">Generic object</typeparam>
    /// <param name="table">DataTable</param>
    /// <param name="options"></param>
    /// <returns>List with generic objects</returns>
    public static IEnumerable<T> AsEnumerable<T>(this DataTable table, ReadOptions options, CancellationToken cancellationToken = default)
        where T : class, new()
    {
        foreach (DataRow row in table.Rows)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var obj = new T();

            foreach (var prop in obj.GetType().GetProperties())
            {
                try
                {
                    PropertyInfo p = obj.GetType().GetProperty(prop.Name);

                    if (p == null) continue;


                    string GetFieldNameFromCustomAttribute()
                    {
                        var atts = p.CustomAttributes.FirstOrDefault(c => c.AttributeType == typeof(ColumnTitleAttribute)
                                                                      || c.AttributeType == typeof(ColumnNameAttribute)
                                                                      || c.AttributeType == typeof(DisplayNameAttribute)
                                                                          );


                        if (options == null || options.TitlesInFirstRow)
                        {
                            //TODO: Refact when remove ColumnTitleAttribute Support

                            if (atts == null)
                            {
                                return p.Name;
                            }

                            if (!string.IsNullOrEmpty(p.GetCustomAttribute<DisplayNameAttribute>()?.DisplayName))
                            {
                                return p.GetCustomAttribute<DisplayNameAttribute>().DisplayName;
                            }

                            return !string.IsNullOrEmpty(p.GetCustomAttribute<ColumnTitleAttribute>()?.DisplayName)
                                ? p.GetCustomAttribute<ColumnTitleAttribute>().DisplayName
                                : p.Name;
                        }
                        else
                        {
                            return atts != null && !string.IsNullOrEmpty(p.GetCustomAttribute<ColumnNameAttribute>()?.ColumnName)
                                ? p.GetCustomAttribute<ColumnNameAttribute>().ColumnName
                                : p.Name;
                        }
                    }


                    var fieldName = GetFieldNameFromCustomAttribute();

                    object ChangeType(object value, Type t)
                    {
                        if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
                        {
                            if (value == null) return null;
                            t = Nullable.GetUnderlyingType(t);
                        }

                        return Convert.ChangeType(value, (Type)t);
                    }


                    //parseamos las celdas vacías que no contienen cadenas
                    var objValue = row[fieldName];
                    if (objValue == System.DBNull.Value) { objValue = string.Empty; }

                    //dejamos valores por defecto en caso de no poder obtener el valor. Así mejor soporte a los nulables
                    if (string.IsNullOrEmpty(objValue as string) && p.PropertyType != typeof(string)) continue;



                    //Converters to parse data to typed field
                    if (options?.Converters != null && options.Converters.Any() && options.Converters.ContainsKey(p.Name))
                    {
                        var exp = options.Converters[p.Name];
                        var f = exp.Compile();
                        p.SetValue(obj, ChangeType(f.DynamicInvoke(objValue), p.PropertyType), null);
                    }
                    else
                    {
                        p.SetValue(obj, ChangeType(objValue, p.PropertyType), null);
                    }


                }
                catch (Exception ex)
                {
                    if (System.Diagnostics.Debugger.IsAttached)
                    {
                        PropertyInfo p = obj.GetType().GetProperty(prop.Name);

                        Console.WriteLine($"{ex.Message} {p.Name}");
                    }

                    continue;
                }
            }

            yield return obj;

            // list.Add(obj);
        }
    }

    #endregion
}

