using System;

namespace BalNET.Infraestructure.ExcelReader.Attributes
{
    [System.AttributeUsage(System.AttributeTargets.Property)]
    public class TableTittleAttribute : Attribute
    {
        public string Title { get; set; }

        public TableTittleAttribute()
        {
            Title = string.Empty;
        }

        public TableTittleAttribute(string title)
        {
            Title = title;
        }
    }
}
