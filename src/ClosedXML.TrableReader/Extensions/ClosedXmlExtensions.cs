using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

// ReSharper disable once CheckNamespace
namespace ClosedXML.Excel
{
    public static class ClosedXmlExtensions
    {
        /// <summary>
        /// Obtiene el contenido de una celda evitando los espacios
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="valueIfNull"></param>
        /// <returns></returns>
        public static string GetContentWithOutSpaces(this IXLCell cell, string valueIfNull = "")
        {
            if (cell.Value == null) return valueIfNull;

            return cell.Value.ToString().Replace(" ", string.Empty);
        }

        /// <summary>
        /// Obtiene un nombre válido para utilizar de nombre de un dataTable limitandolo a 50 cars
        /// </summary>
        /// <param name="ws">IXLWorksheet</param>
        /// <param name="maxNameLength"></param>
        /// <returns></returns>
        public static string GetNameForDataTable(this IXLWorksheet ws,int maxNameLength=50)
        {
            string aux = ws.Name.Replace(" ", "_");

            return aux.Length > maxNameLength ? aux.Substring(0, maxNameLength) : aux;
        }
    }
}
