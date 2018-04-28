using System;
using ClosedXML.TableReader.Attributes;

namespace Net.Core.Sample.Models
{
    public class TableWithHeadersAndBoleanConversion
    {
        public int Number { get; set; }
        public string Name { get; set; }

        [ColumnTittle("Birthday date")]
        public DateTime Birthday { get; set; }

        public DateTime ADate { get; set; }


        public bool Selected { get; set; }

    }
}