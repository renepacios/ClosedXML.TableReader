// ReSharper disable once CheckNamespace
namespace ClosedXML.Excel;

public static class ClosedXmlExtensions
{
    /// <summary>
    /// Get cell content without spaces
    /// </summary>
    /// <param name="cell">XLCell instance <see cref="IXLCell"/></param>
    /// <param name="valueIfNull"> Set a default value to empty cells </param>
    /// <returns></returns>
    public static string GetContentWithOutSpaces(this IXLCell cell, string valueIfNull = "")
    {
        string contentWithout = valueIfNull;

        if (cell.HasFormula && !cell.CachedValue.IsBlank)
        {
            contentWithout = cell.CachedValue.ToString().Replace(" ", string.Empty);
        }
        else
        {
            contentWithout = cell.Value.ToString().Replace(" ", string.Empty);
        }

        return contentWithout;

    }

    /// <summary>
    /// Get cell content remove the without spaces at the beginning and end
    /// </summary>
    /// <param name="cell">XLCell instance <see cref="IXLCell"/></param>
    /// <param name="valueIfNull"> Set a default value to empty cells </param>
    /// <returns></returns>
    public static string GetContentTrim(this IXLCell cell, string valueIfNull = "")
    {
        if (cell.Value.IsBlank) return valueIfNull;

        return cell.Value.ToString().Trim();
    }

    /// <summary>
    /// Get cell content in Type T
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="cell">XLCell instance <see cref="IXLCell"/></param>
    /// <param name="conversor">Func convert types</param>
    /// <param name="valueIfNull"> Default value return to empty cells</param>
    /// <returns></returns>
    public static T GetContentTyped<T>(this IXLCell cell, Func<string, T> conversor, T valueIfNull = default(T))
    {
        string valor = cell.GetContentWithOutSpaces(null);
        if (string.IsNullOrEmpty(valor)) return default(T);

        return conversor(valor);
    }


    /// <summary>
    /// Get Worksheet name without spaces an limit it lenght
    /// </summary>
    /// <param name="ws">IXLWorksheet instance <see cref="IXLWorksheet"/></param>
    /// <param name="maxNameLength">Name max length</param>
    /// <returns></returns>
    public static string GetNameForDataTable(this IXLWorksheet ws, int maxNameLength = 50)
    {
        string aux = ws.Name.Replace(" ", "_");

        return aux.Length > maxNameLength ? aux.Substring(0, maxNameLength) : aux;
    }
}
