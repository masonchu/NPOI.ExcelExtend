using System;

namespace NPOI.ExcelExtend.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        public short Order { get; set; }
        public string Format { get; set; }

        public ExcelColumnAttribute()
        {
            Order = short.MaxValue;
            Format = string.Empty;
        }
    }
}
