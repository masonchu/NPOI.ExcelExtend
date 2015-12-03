using System;

namespace NPOI.ExcelExport.Attributes
{
    public class ChildDataAttribute : Attribute
    {
        public Type type { get; set; }
        public ChildDataAttribute(Type type)
        {
            this.type = type;
        }
    }
}
