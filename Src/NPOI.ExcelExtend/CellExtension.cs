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
        public static void SetCellValueByType(this ICell cell, IWorkbook workbook, object model, string dataFormat = null)
        {
            double number;
            DateTime datetime;
            bool modelBool;
            var newDataFormat = workbook.CreateDataFormat();
            var style = workbook.CreateCellStyle();
            cell.CellStyle = SetStyle(workbook, dataFormat);


            if (TryParseNumeric(model, out number))
                cell.SetCellValue(number);
            else if (TryParseBoolean(model, out modelBool))
                cell.SetCellValue(modelBool);
            else if (TryParseDateTime(model, out datetime))
            {
                cell.CellStyle = SetStyle(workbook, "yyyy/MM/dd HH:mm:ss");

                cell.SetCellValue(datetime);
            }
            else
            {
                cell.SetCellValue(model.ToString());
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

            return Double.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), System.Globalization.NumberStyles.Any, NumberFormatInfo.InvariantInfo, out result);
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

            return DateTime.TryParse(Convert.ToString(expression, CultureInfo.InvariantCulture), out result);
        }
    }
}


