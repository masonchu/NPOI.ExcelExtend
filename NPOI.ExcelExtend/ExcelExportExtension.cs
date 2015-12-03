using NPOI.ExcelExport.Attributes;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Reflection;

namespace NPOI.ExcelExport
{
    public static class ExportExcelExtensions
    {
        public static MemoryStream ExportExcel<T>(this IEnumerable<T> dataList)
        {
            //Create workbook
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet(string.Format("{0}", "Sheet1"));
            dataList.ExcelSheet(worksheet);

            MemoryStream sw = new MemoryStream();

            workbook.Write(sw);
            return sw;

            //Save file
            //FileStream file = new FileStream(fileName, FileMode.Create);
            //workbook.Write(file);
            //file.Close();
        }

        /// <summary>
        /// edit this sheet
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataList"></param>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static ISheet ExcelSheet<T>(this IEnumerable<T> dataList, ISheet worksheet)
        {
            var datatype = typeof(T);

            //Insert titles
            var row = worksheet.CreateRow(0);
            var titleList = datatype.GetPropertyDisplayNames();
            for (int cellNumber = 0; cellNumber < titleList.Count; cellNumber++)
            {
                row.CreateCell(cellNumber).SetCellValue(titleList[cellNumber]);
            }

            var numberOfColumns = 0;
            //Insert data values
            var rowNumber = 1;
            foreach (var row_item in dataList)
            {
                ///row
                var valueList = row_item.GetPropertyValues();
                //var mergeRows = item.GetPropertyMergeRows();
                ///TODO if need merge columns
                //if (mergeRows > 1)
                //{
                //    worksheet.AddMergedRegion(new CellRangeAddress(rowNumber, rowNumber + mergeRows, rowNumber, rowNumber));
                //}
                var cellNumber = 0;
                var new_rowData = true;
                foreach (var values in valueList)
                {
                    var tmpRow = worksheet.GetRow(rowNumber);
                    if (tmpRow == null)
                    {
                        tmpRow = worksheet.CreateRow(rowNumber);
                    }


                    foreach (var cell in values)
                    {
                        tmpRow.CreateCell(cellNumber).SetCellValue(cell);
                        numberOfColumns = cellNumber;
                        cellNumber++;
                    }
                    if (new_rowData)
                    {
                        new_rowData = false;
                    }
                    else
                    {
                        cellNumber = cellNumber - values.Count;
                        rowNumber++;
                    }


                }
                rowNumber++;
            }

            worksheet.Autobreaks = true;
            for (int i = 0; i <= numberOfColumns; i++)
            {
                worksheet.AutoSizeColumn(i);
            }

            return worksheet;
        }

        public static void ExportExcel<T>(this IEnumerable<T> dataList, string fileName)
        {
            //Create workbook
            var datatype = typeof(T);
            var workbook = new XSSFWorkbook();
            var worksheet = workbook.CreateSheet(string.Format("{0}", datatype.GetDisplayName()));

            dataList.ExcelSheet(worksheet);

            //Save file
            FileStream file = new FileStream(fileName, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }
        /// <summary>
        /// get display names
        /// </summary>
        /// <param name="memberInfo"></param>
        /// <returns></returns>
        public static string GetDisplayName(this MemberInfo memberInfo)
        {
            var titleName = string.Empty;

            //Try get DisplayName
            var attribute = memberInfo.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();
            if (attribute != null)
            {
                titleName = (attribute as DisplayAttribute).Name;
            }
            //If no DisplayName
            else
            {
                titleName = memberInfo.Name;
            }

            return titleName;
        }

        /// <summary>
        /// get headers
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<string> GetPropertyDisplayNames(this Type type, List<string> titleList = null)
        {
            titleList = (titleList == null ? new List<string>() : titleList);
            var propertyInfos = type.GetProperties();

            foreach (var propertyInfo in propertyInfos)
            {
                object[] attrs = propertyInfo.GetCustomAttributes(true);
                foreach (object attr in attrs)
                {
                    ///if this column need to export
                    ExcelColumnAttribute authAttr = attr as ExcelColumnAttribute;
                    if (authAttr != null)
                    {
                        var titleName = propertyInfo.GetDisplayName();

                        titleList.Add(titleName);
                    }

                    ///if this object of collection need to export
                    var childDataAttr = attr as ChildDataAttribute;
                    if (childDataAttr != null)
                    {
                        GetPropertyDisplayNames(childDataAttr.type, titleList);
                    }
                }

            }

            return titleList;
        }
        /// <summary>
        /// get a martix of object 
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static List<List<string>> GetPropertyValues<T>(this T data, List<List<string>> allValues = null)
        {
            allValues = allValues == null ? new List<List<string>>() : allValues;
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
                    ///if this object of collection need to export
                    var childDataAttr = attr as ChildDataAttribute;
                    if (childDataAttr != null)
                    {
                        var vals = propertyInfo.GetValue(data, null);
                        var val = vals as IEnumerable<Object>;
                        foreach (var item in val)
                        {
                            GetPropertyValues(item, allValues);
                        }
                        //GetPropertyDisplayNames()
                    }
                }
            }
            allValues.Insert(0, propertyValues);

            return allValues;
        }
    }
}
