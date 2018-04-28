using System;
using System.Collections.Generic;
using System.Linq.Expressions;

namespace ClosedXML.TableReader.Model
{
    public class ReadOptions
    {

        public ReadOptions()
        {
            Converters = new Dictionary<string, LambdaExpression>();
        }

        /// <summary>
        /// Genera opciones por defecto si no se indican al consumir la librería
        /// </summary>
        public static ReadOptions DefaultOptions => new ReadOptions()
        {
            TitlesInFirstRow = true
        };

        public bool TitlesInFirstRow { get; set; }

        public Dictionary<string, LambdaExpression> Converters { get; set; }
    }





}