using NPOI.ExcelExtend.Test.Models;
using NPOI.ExcelExtend.Test.Models.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOI.ExcelExtend.Test.Respostories
{
    public class MemberRespostory
    {
        public IQueryable<Member> All()
        {
            return new Member[] {
                new Member() { FirstName = "Mason",LastName="Chu",Age =30,Gender = Gender.Male ,Height=178.21m, UpdateOn=DateTime.Now,IsMarried=false }
            }.AsQueryable();
        }
    }
}
