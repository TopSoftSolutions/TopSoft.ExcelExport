using System;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TopSoft.ExcelExport.Helpers
{
    static class ExcelHelper
    {
        public static int ColumnCompare(string strA, string strB)
        {
            int retVal = 0;

            Regex charPartRegex = new Regex(@"[^0-9]*", RegexOptions.IgnoreCase);

            Match charPartStrAMatch = charPartRegex.Match(strA);
            Match charPartStrBMatch = charPartRegex.Match(strB);

            string charPartStrA = charPartStrAMatch.Groups[0].Value;
            string charPartStrB = charPartStrBMatch.Groups[0].Value;

            int charPartLengthStrA = charPartStrA.Count();
            int charPartLengthStrB = charPartStrB.Count();

            if(charPartLengthStrA > charPartLengthStrB)
            {
                return 1;
            }
            else if(charPartLengthStrA < charPartLengthStrB)
            {
                return -1;
            }

            //     A 32-bit signed integer that indicates the lexical relationship between the
            //     two comparands.Value Condition Less than zero strA is less than strB. Zero
            //     strA equals strB. Greater than zero strA is greater than strB.
            retVal = string.Compare(charPartStrA, charPartStrB, true);

            if(retVal == 0)
            {
                int numberPartStrA = Convert.ToInt32(strA.Replace(charPartStrA, string.Empty));
                int numberPartStrB = Convert.ToInt32(strB.Replace(charPartStrB, string.Empty));
                if(numberPartStrA < numberPartStrB)
                {
                    retVal = -1;
                }
                else if(numberPartStrA > numberPartStrB)
                {
                    retVal = 1;
                }
            }

            return retVal;
        }

        public static CellValues ResolveCellType(Type propertyType)
        {
            var retCellType = CellValues.String;

            if(propertyType == typeof(string))
            {
                retCellType = CellValues.String;
            }
            else if(propertyType == typeof(sbyte) || propertyType == typeof(byte) || propertyType == typeof(short) || propertyType == typeof(ushort)
                    || propertyType == typeof(int) || propertyType == typeof(uint) || propertyType == typeof(long) || propertyType == typeof(ulong)
                    || propertyType == typeof(float) || propertyType == typeof(double) || propertyType == typeof(decimal))
            {
                retCellType = CellValues.Number;
            }
            else if(propertyType == typeof(bool))
            {
                retCellType = CellValues.Boolean;
            }
            else
            {
                throw new Exception(string.Format("Unsuported property of type {0} in ExcelRow.", propertyType.Name));
            }

            return retCellType;
        }
    }
}
