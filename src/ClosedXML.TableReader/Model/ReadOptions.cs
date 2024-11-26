using System.Linq.Expressions;

namespace ClosedXML.TableReader.Model;

public class ReadOptions
{
    /// <summary> Genera opciones por defecto si no se indican al consumir la librería </summary>
    public static ReadOptions DefaultOptions => new()
    {
        TitlesInFirstRow = true,
        RowStart = 0,
    };

    public bool TitlesInFirstRow { get; set; }

    /// <summary> number row of content table </summary>
    public int RowStart { get; set; }


    // /// <summary> number row of content table </summary>
    // public int ColumnStart { get; set; }

    public Dictionary<string, LambdaExpression> Converters { get; set; } = new();
}