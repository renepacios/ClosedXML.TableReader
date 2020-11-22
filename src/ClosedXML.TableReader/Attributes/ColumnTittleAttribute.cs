using System;

namespace ClosedXML.TableReader.Attributes
{
    /// <summary>
    /// Pair property value with data in column with this title
    /// <remarks>
    ///     Use only with ReadOptions.TitlesInFirstRow=true <see cref="ClosedXML.TableReader.Model.ReadOptions"/>
    /// </remarks>
    /// </summary>
    [System.AttributeUsage(System.AttributeTargets.Property)]
    //[Obsolete("Use DisplayName")]
    internal class ColumnTittleAttribute : Attribute
    {
        public string Title { get; set; }

        public ColumnTittleAttribute()
        {
            Title = string.Empty;
        }

        public ColumnTittleAttribute(string title)
        {
            Title = title;
        }
    }
}
