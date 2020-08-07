namespace ClassicWinFormsApplication.Models
{
    using System;
    using System.ComponentModel;

    public class TableWithHeaders
    {
        public int Number { get; set; } 
        public string Name { get; set; }

        [DisplayName("Birthday date")]
        public DateTime Birthday { get; set; }

        public DateTime ADate { get; set; }

    }
}
