

// ReSharper disable once CheckNamespace
namespace ClosedXML.Excel
{
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
            if (cell.Value == null) return valueIfNull;

            return cell.Value.ToString().Replace(" ", string.Empty);
        }

        /// <summary>
        /// Get Worksheet name without spaces an limit it lenght
        /// </summary>
        /// <param name="ws">IXLWorksheet instance <see cref="IXLWorksheet"/></param>
        /// <param name="maxNameLength">Name max length</param>
        /// <returns></returns>
        public static string GetNameForDataTable(this IXLWorksheet ws,int maxNameLength=50)
        {
            string aux = ws.Name.Replace(" ", "_");

            return aux.Length > maxNameLength ? aux.Substring(0, maxNameLength) : aux;
        }
    }
}
