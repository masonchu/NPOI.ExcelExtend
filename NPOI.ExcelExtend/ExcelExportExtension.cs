
using NPOI.ExcelExport;
using NPOI.ExcelExport.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NAWF.Utilities.Extension
{
    public static class ExportExcelExtensions
    {
        public static MemoryStream ExportExcel<T>(this IEnumerable<T> dataList)
        {
            //Create workbook
            var datatype = typeof(T);
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet(string.Format("{0}", "Sheet1"));

            //Insert titles
            var row = worksheet.CreateRow(0);
            var titleList = datatype.GetPropertyDisplayNames();
            for (int i = 0; i < titleList.Count; i++)
            {
                row.CreateCell(i).SetCellValue(titleList[i]);
            }

            var numberOfColumns = 0;
            //Insert data values
            for (int i = 1; i < dataList.Count() + 1; i++)
            {
                var tmpRow = worksheet.CreateRow(i);
                var valueList = dataList.ElementAt(i - 1).GetPropertyValues();

                for (int j = 0; j < valueList.Count; j++)
                {
                    tmpRow.CreateCell(j).SetCellValue(valueList[j]);
                    numberOfColumns = j;
                }
            }

            worksheet.Autobreaks = true;
            for (int i = 0; i <= numberOfColumns; i++)
            {
                worksheet.AutoSizeColumn(i);
            }
            MemoryStream sw = new MemoryStream();

            workbook.Write(sw);
            return sw;

            //Save file
            //FileStream file = new FileStream(fileName, FileMode.Create);
            //workbook.Write(file);

            //file.Close();
        }
        public static void ExportExcel<T>(this IEnumerable<T> dataList, string fileName)
        {
            //Create workbook
            var datatype = typeof(T);
            var workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet(string.Format("{0}", datatype.GetDisplayName()));

            //Insert titles
            var row = worksheet.CreateRow(0);
            var titleList = datatype.GetPropertyDisplayNames();
            for (int i = 0; i < titleList.Count; i++)
            {
                row.CreateCell(i).SetCellValue(titleList[i]);
            }

            //Insert data values
            for (int i = 1; i < dataList.Count() + 1; i++)
            {
                var tmpRow = worksheet.CreateRow(i);
                var valueList = dataList.ElementAt(i - 1).GetPropertyValues();

                for (int j = 0; j < valueList.Count; j++)
                {
                    tmpRow.CreateCell(j).SetCellValue(valueList[j]);
                }
            }

            //Save file
            FileStream file = new FileStream(fileName, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }

        public static string GetDisplayName(this MemberInfo memberInfo)
        {
            var titleName = string.Empty;

            //Try get DisplayName
            var attribute = memberInfo.GetCustomAttributes(typeof(DisplayNameAttribute), false).FirstOrDefault();
            if (attribute != null)
            {
                titleName = (attribute as DisplayNameAttribute).DisplayName;
            }
            //If no DisplayName
            else
            {
                titleName = memberInfo.Name;
            }

            return titleName;
        }

        public static List<string> GetPropertyDisplayNames(this Type type)
        {
            var titleList = new List<string>();
            var propertyInfos = type.GetProperties();

            foreach (var propertyInfo in propertyInfos)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        var titleName = propertyInfo.GetDisplayName();

                        titleList.Add(titleName);
                    }
                }

            }

            return titleList;
        }

        public static List<string> GetPropertyValues<T>(this T data)
        {
            var propertyValues = new List<string>();
            var propertyInfos = data.GetType().GetProperties();

            foreach (var propertyInfo in propertyInfos)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        var val = propertyInfo.GetValue(data, null);
                        var valString = val == null ? "" : val.ToString();
                        propertyValues.Add(valString);
                    }
                }
            }

            return propertyValues;
        }
    }
}
