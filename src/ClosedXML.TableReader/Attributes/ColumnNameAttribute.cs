namespace ClosedXML.TableReader.Attributes;

/// <summary>
/// Pair property value with data in column with this name
/// <example>
///     [ColumnName("A")] ,[ColumnName("AC")]
/// </example>
/// <remarks>
///     Use only with ReadOptions.TitlesInFirstRow=false <see cref="ClosedXML.TableReader.Model.ReadOptions"/>
/// </remarks>
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ColumnNameAttribute(string columnName = "") : Attribute
{
    public string ColumnName { get; set; } = columnName;

}
