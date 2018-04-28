using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.TableReader;
using ClosedXML.TableReader.Attributes;

namespace Net.Core.Sample.Models
{
    public class TableWithHeaders
    {
        public int Number { get; set; } 
        public string Name { get; set; }

        [ColumnTittle("Birthday date")]
        public DateTime Birthday { get; set; }

        public DateTime ADate { get; set; }

    }
}
