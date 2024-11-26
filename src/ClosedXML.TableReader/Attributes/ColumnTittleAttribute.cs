using System;

namespace ClosedXML.TableReader.Attributes;

/// <summary>
/// Pair property value with data in column with this title
/// <remarks>
///     Use only with ReadOptions.TitlesInFirstRow=true <see cref="ClosedXML.TableReader.Model.ReadOptions"/>
/// </remarks>
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
//[Obsolete("Use DisplayName")]
internal class ColumnTittleAttribute : Attribute
{
    public string Title { get; set; } = string.Empty;

    public ColumnTittleAttribute(string title) => Title = title;
}
