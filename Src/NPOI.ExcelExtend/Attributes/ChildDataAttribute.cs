using System;

namespace NPOI.ExcelExtend.Attributes
{
    /// <summary>
    /// Export ICollection with this attribute
    /// </summary>
    public class ChildDataAttribute : Attribute
    {
        public Type type { get; set; }
        public ChildDataAttribute(Type type)
        {
            this.type = type;
        }
    }
}
