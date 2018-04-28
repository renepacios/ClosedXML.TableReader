using System;
using ClosedXML.TableReader.Model;

namespace ClosedXML.TableReader.Attributes
{
    /// <summary>
    /// Pair property value with data in column with this name
    /// <example>
    ///     [ColumnName("A")] ,[ColumnName("AC")]
    /// </example>
    /// <remarks>
    ///     Use only with ReadOptions.TitlesInFirstRow=false <see cref="ClosedXML.TableReader.Model.ReadOptions"/>
    /// </remarks>
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class ColumnNameAttribute : Attribute
    {
        public string ColumnName { get; set; }

        public ColumnNameAttribute()
        {
            ColumnName = string.Empty;
        }

        public ColumnNameAttribute(string columnName)
        {
            ColumnName = columnName;
        }
    }
}
