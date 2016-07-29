using NPOI.ExcelExtend.Attributes;
using NPOI.ExcelExtend.Test.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.ExcelExtend.Test.Models
{
    public class Member
    {
        [ExcelColumn(Order = 1)]
        public string FirstName { get; set; }
        [ExcelColumn(Order = 2)]
        public string LastName { get; set; }
        [ExcelColumn]
        public bool IsMarried { get; set; }

        [ExcelColumn]
        public DateTime UpdateOn { get; set; }
        [ExcelColumn(Order = 5)]
        public int Age { get; set; }
        [ExcelColumn(Order = 4)]
        public decimal Height { get; set; }
        [ExcelColumn(Order = 3)]
        public Gender Gender { get; set; }
    }
}
