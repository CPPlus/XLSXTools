using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXTools
{
    public class XLSXUtils
    {
        public static int CellReferenceToRowIndex(string cellReference)
        {
            return int.Parse(GetNumbers(cellReference));
        }

        public static int CellReferenceToColumnIndex(string cellReference)
        {
            return LetterIndexToInt(GetLetters(cellReference));
        }

        public static int LetterIndexToInt(string letterIndex)
        {
            letterIndex = letterIndex.ToUpperInvariant();

            int sum = 0;
            for (int i = 0; i < letterIndex.Length; i++)
            {
                sum *= 26;
                sum += (letterIndex[i] - 'A' + 1);
            }

            return sum;
        }

        public static string ColumnIndexToLetter(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        private static string GetLetters(string text)
        {
            return new string(text.Where(Char.IsLetter).ToArray());
        }

        private static string GetNumbers(string text)
        {
            return new string(text.Where(Char.IsDigit).ToArray());
        }
    }
}
