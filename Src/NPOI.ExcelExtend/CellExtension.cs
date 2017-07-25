using NPOI.ExcelExtend.Models;
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
        public static void SetCellValueByType(this ICell cell, IWorkbook workbook, CellModel model)
        {
            if (model != null && model.Value != null)
            {
                double number;
                DateTime datetime;
                bool modelBool;

                if (!string.IsNullOrWhiteSpace(model.Format))
                {
                    cell.CellStyle = GetStyle(workbook, model.Format);
                }

                if (TryParseNumeric(model.Value, out number))
                {
                    cell.SetCellValue(number);
                }
                else if (TryParseBoolean(model.Value, out modelBool))
                {
                    cell.SetCellValue(modelBool);
                }
                else if (TryParseDateTime(model.Value, out datetime))
                {
                    if (string.IsNullOrWhiteSpace(model.Format))
                    {
                        cell.CellStyle = GetStyle(workbook, "yyyy/MM/dd HH:mm:ss");
                    }
                    cell.SetCellValue(datetime);
                }
                else
                {
                    cell.SetCellType(CellType.String);
                    cell.SetCellValue(model.Value.ToString());
                }
            }

        }

        public static ICellStyle GetStyle(IWorkbook workbook, short datafmt)
        {
            for (short i = 0; i < workbook.NumCellStyles; i++)
            {
                var style = workbook.GetCellStyleAt(i);
                if (style.DataFormat == datafmt) return style;
            }
            return null;
        }

        public static ICellStyle GetStyle(IWorkbook workbook, string value)
        {
            var datafmt = workbook.CreateDataFormat().GetFormat(value);
            var cs = GetStyle(workbook, datafmt);
            if (cs == null)
            {
                cs = workbook.CreateCellStyle();
                cs.DataFormat = datafmt;
            }
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


