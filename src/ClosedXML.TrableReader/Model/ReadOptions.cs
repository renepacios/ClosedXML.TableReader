namespace ClosedXML.TableReader.Model
{
    public class ReadOptions
    {
        /// <summary>
        /// Genera opciones por defecto si no se indican al consumir la librería
        /// </summary>
        public static ReadOptions DefaultOptions => new ReadOptions()
        {
            TitlesInFirstRow = true
        };

        public bool TitlesInFirstRow { get; set; }
    }
}