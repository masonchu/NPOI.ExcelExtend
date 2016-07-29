using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace NPOI.ExcelExtend.Test
{
    public class BaseUnitTest
    {
        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        protected void OuputObjectJson(object model)
        {
            try
            {
                var json = JsonConvert.SerializeObject(model, Formatting.Indented)
                    .Replace("{", "<").Replace("}", ">");
                TestContext.WriteLine(json);
            }
            catch (Exception ex)
            {
                TestContext.WriteLine("Debug Output Error : {0}", ex.Message);
            }
        }
    }
}
