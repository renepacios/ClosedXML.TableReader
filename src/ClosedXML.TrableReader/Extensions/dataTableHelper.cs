using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using ClosedXML.TableReader.Attributes;

namespace ClosedXML.TableReader.Extensions
{
    public static class DataTableHelper
    {
        /// <summary>
        /// Converts a DataTable to a list with generic objects
        /// </summary>
        /// <typeparam name="T">Generic object</typeparam>
        /// <param name="table">DataTable</param>
        /// <returns>List with generic objects</returns>
        public static IEnumerable<T> AsEnumerableTyped<T>(this DataTable table) where T : class, new()
        {
           
                //List<T> list = new List<T>();

                foreach (var row in table.AsEnumerable())
                {
                    T obj = new T();

                    foreach (var prop in obj.GetType().GetProperties())
                    {
                        try
                        {
                            PropertyInfo p = obj.GetType().GetProperty(prop.Name);

                            if (p == null) continue;

                            var atts = p.CustomAttributes.FirstOrDefault(c => c.AttributeType==typeof(TableTittleAttribute));

                            
                            var fieldName = atts != null && !string.IsNullOrEmpty(p.GetCustomAttribute<TableTittleAttribute>().Title)
                                ? p.GetCustomAttribute<TableTittleAttribute>().Title
                                : p.Name;

                            p.SetValue(obj,
                                Convert.ChangeType(row[fieldName], p.PropertyType), null);
                        }
                        catch
                        {
                            continue;
                        }
                    }

                    yield return obj;

                    // list.Add(obj);
                }

        }
    }
}

