using System.ComponentModel;

namespace ClosedXML.TableReader.Attributes;

/// <summary>
/// Pair property value with data in column with this title
/// <remarks>
///     Use only with ReadOptions.TitlesInFirstRow=true <see cref="ClosedXML.TableReader.Model.ReadOptions"/>
/// </remarks>
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
internal class ColumnTitleAttribute : DisplayNameAttribute
{
    public ColumnTitleAttribute()
    {
    }

    public ColumnTitleAttribute(string displayName) : base(displayName)
    {
    }
}
