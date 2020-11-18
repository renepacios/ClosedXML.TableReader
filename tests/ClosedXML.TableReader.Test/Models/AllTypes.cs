using System;

namespace ClosedXML.TableReader.Test.Models
{
    public class AllTypes
    {
        public string Description { get; set; }

        public int Integer { get; set; }
        public byte Byte { get; set; }
        public char Char { get; set; }
        public string String { get; set; }
        public DateTime Date { get; set; }
        public float Float { get; set; }
        public double Double { get; set; }
        public bool BoleanString { get; set; }

        }

    public class AllTypesWithNulls
    {
        public string Description { get; set; }

        public int? Integer { get; set; }
        public byte? Byte { get; set; }
        public char? Char { get; set; }
        public string String { get; set; }
        public DateTime? Date { get; set; }
        public float? Float { get; set; }
        public double? Double { get; set; }
        public bool? BoleanString { get; set; }

    }
}