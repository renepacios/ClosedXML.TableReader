using System;
using System.ComponentModel;

namespace ClosedXML.TableReader.Test.Models
{
    public class TableWithHeadersAndBoleanConversion
    {
        public int Number { get; set; }
        public string Name { get; set; }

        [DisplayName("Birthday date")]
        public DateTime Birthday { get; set; }

        public DateTime ADate { get; set; }
        
        public bool Selected { get; set; }

    }

    public class TableWithHeadersAndNullableBoleanConversion
    {
        public int Number { get; set; }
        public string Name { get; set; }

        [DisplayName("Birthday date")]
        public DateTime Birthday { get; set; }

        public DateTime ADate { get; set; }

        public bool? Selected { get; set; }
    }

}