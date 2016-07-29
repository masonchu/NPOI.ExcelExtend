using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace NPOI.ExcelExtend
{
    internal static class CellExtension
    {
        public static void SetCellValueByType(this ICell cell, IWorkbook workbook, object model, string dataFormat = "")
        {
            if (model != null)
            {
                double number;
                DateTime datetime;
                bool modelBool;
                var newDataFormat = workbook.CreateDataFormat();
                var style = workbook.CreateCellStyle();
                cell.CellStyle = SetStyle(workbook, dataFormat);


                if (TryParseNumeric(model, out number))
                {
                    cell.SetCellValue(number);
                }
                else if (TryParseBoolean(model, out modelBool))
                    cell.SetCellValue(modelBool);
                else if (TryParseDateTime(model, out datetime))
                {
                    cell.CellStyle = SetStyle(workbook, "yyyy/MM/dd HH:mm:ss");
                    cell.SetCellValue(datetime);
                }
                else
                {
                    var modelString = model.ToString();
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(modelString);
                }
            }

        }

        public static XSSFCellStyle SetStyle(IWorkbook workbook, string value)
        {
            XSSFCellStyle cs = (XSSFCellStyle)workbook.CreateCellStyle();
            //cs.BorderBottom = BorderStyle.Thin;
            //cs.BorderTop = BorderStyle.Thin;
            //cs.BorderLeft = BorderStyle.Thin;
            //cs.BorderRight = BorderStyle.Thin;
            //XSSFFont font = (XSSFFont)workbook.CreateFont();
            //font.FontHeightInPoints = 12;
            XSSFDataFormat format = (XSSFDataFormat)workbook.CreateDataFormat();
            //cs.SetFont(font);
            if (!string.IsNullOrWhiteSpace(value))
                cs.DataFormat = format.GetFormat(value);
            return cs;
        }

        private static bool TryParseNumeric(object expression, out double result)
        {
            //if (expression == null)
            //    return false;
            if (expression.IsNumeric())
            {
                return Double.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Any, NumberFormatInfo.InvariantInfo, out result);
            }
            else
            {
                result = 0;
                return false;
            }
        }

        private static bool TryParseBoolean(object expression, out bool result)
        {
            //if (expression == null)
            //    return false;

            return Boolean.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), out result);
        }

        private static bool TryParseDateTime(object expression, out DateTime result)
        {
            //if (expression == null)
            //    return false;
            if (expression.IsDateTime())
            {
                return DateTime.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), out result);
            }
            else
            {
                result = DateTime.Now;
                return false;
            }
        }

        private static bool IsNumeric(this object value)
        {
            return value is sbyte
                    || value is byte
                    || value is short
                    || value is ushort
                    || value is int
                    || value is uint
                    || value is long
                    || value is ulong
                    || value is float
                    || value is double
                    || value is decimal;
        }

        private static bool IsDateTime(this object value)
        {
            return value is DateTime;
        }
    }
}


