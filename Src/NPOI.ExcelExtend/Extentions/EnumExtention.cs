using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace NPOI.ExcelExtend.Extentions
{
    public static class EnumExtention
    {
        /// <summary>
        /// get enum description
        /// </summary>
        /// <param name="enum"></param>
        /// <returns></returns>
        public static string ToDescriptionString(this object @enum)
        {
            var attribute = @enum.GetType().GetField(@enum.ToString()).GetCustomAttributes(typeof(DescriptionAttribute), true).FirstOrDefault();
            return attribute == null ? @enum.ToString() : ((DescriptionAttribute)attribute).Description;
        }

        public static string ToDescriptionString(this System.Enum @enum, string enumname)
        {
            try
            {
                var attribute = @enum.GetType().GetField(enumname).GetCustomAttributes(typeof(DescriptionAttribute), true).FirstOrDefault();
                return attribute == null ? @enum.ToString() : ((DescriptionAttribute)attribute).Description;
            }
            catch
            {
                return null;
            }
        }
    }
}
