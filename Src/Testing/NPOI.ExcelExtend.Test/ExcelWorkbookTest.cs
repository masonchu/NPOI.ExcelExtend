using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.ExcelExtend.Test.Respostories;
using System.Linq;
using NPOI.ExcelExtend;
using System.IO;

namespace NPOI.ExcelExtend.Test
{
    /// <summary>
    /// Summary description for ExcelWorkbookTest
    /// </summary>
    [TestClass]
    public class ExcelWorkbookTest : BaseUnitTest
    {
        public ExcelWorkbookTest()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        [TestMethod]
        public void Test_Member_To_ExcelSheet()
        {
            var respostory = new MemberRespostory();
            var target = respostory.All();
            //var actul = target.AsEnumerable().ExportExcel();

            // convert string to stream
            //byte[] buffer = Encoding.ASCII.GetBytes(memString);
            //MemoryStream ms = new MemoryStream(buffer);
            //write to file
            var actul = target.ToList();
            var fileName = string.Format("Member_To_ExcelSheet-{0}.xlsx", Guid.NewGuid().ToString());
            File.WriteAllBytes(fileName, actul.ExportExcel().ToArray());
            //SaveFileStream("file.xlsx", actul);
            OuputObjectJson(fileName);
        }

        //private void SaveFileStream(String path, Stream stream)
        //{
        //    var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write);
        //    stream.CopyTo(fileStream);
        //    fileStream.Dispose();
        //}
    }
}
