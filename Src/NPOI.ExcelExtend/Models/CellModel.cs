using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NPOI.ExcelExtend.Models
{
    public class CellModel
    {
        public object Value { get; set; }
        public string Format { get; set; }
        public short Order { get; set; }
    }
}
