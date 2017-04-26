using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXTools
{
    public class XLSXRowReader
    {
        private int rowIndexToRead = 0;
        private int currentRowIndex = 0;
        private int currentColumnIndex = 0;
        private string currentCellValue = null;
        private bool couldReadLast = true;
        private bool customRangeCalculation;
        private int rowStart;

        private int rowIndexToReadTo = 0;
        private int columnIndexToReadTo = 0;

        private string srcFilePath;
        private string sheetName;
        private XLSXReader reader;

        public bool EOF
        {
            get
            {
                return reader.EOF;
            }
        }

        public int RowCount
        {
            get
            {
                return reader.RowCount;
            }
        }

        public int ColumnCount
        {
            get
            {
                return reader.ColumnCount;
            }
        }

        public XLSXRowReader(string srcFilePath, bool customRangeCalculation = false, int rowStart = 0) : this(srcFilePath, "Sheet1", customRangeCalculation, rowStart)
        {

        }

        public XLSXRowReader(string srcFilePath, string sheetName, bool customRangeCalculation = false, int rowStart = 0)
        {
            this.srcFilePath = srcFilePath;
            this.sheetName = sheetName;
            this.customRangeCalculation = customRangeCalculation;
            this.rowStart = rowStart;

            ResetState();
        }

        public void Close()
        {
            reader.Close();
        }

        public bool ReadNextRecord(out string[] record)
        {
            record = new string[columnIndexToReadTo];
            Cell[] cells;
            bool result = ReadNextCells(out cells);
            if (result)
            {
                for (int i = 0; i < record.Length; i++)
                {
                    record[i] = reader.GetCellValue(cells[i]);
                }
                return true;
            }
            else return false;
        }

        public string GetCellValue(Cell cell)
        {
            return reader.GetCellValue(cell);
        }

        public string GetCellFormat(Cell cell)
        {
            return reader.GetCellFormat(cell);
        }

        public bool ReadNextCells(out Cell[] cells)
        {
            cells = new Cell[columnIndexToReadTo];
            rowIndexToRead++;

            if (!couldReadLast || rowIndexToRead > rowIndexToReadTo) return false;

            do
            {
                UpdateCellsState(reader.CurrentCell);

                if (currentRowIndex == rowIndexToRead)
                {
                    if (currentColumnIndex - 1 < columnIndexToReadTo)
                    {
                        cells[currentColumnIndex - 1] = reader.CurrentCell;
                    }
                }
                else break;
            } while (couldReadLast = reader.ReadNextCell());

            return true;
        }

        public bool IsValidDateTimeFormat(string format)
        {
            return reader.IsValidDateTimeFormat(format);
        }

        private void FillWithEmptyRecord(out string[] record)
        {
            record = new string[columnIndexToReadTo];
            for (int i = 0; i < record.Length; i++)
            {
                record[i] = string.Empty;
            }
        }

        private void ResetState()
        {
            reader = new XLSXReader(srcFilePath, sheetName, customRangeCalculation, rowStart);
            reader.ReadNextCell();

            currentRowIndex = 0;
            currentColumnIndex = 0;

            rowIndexToReadTo = reader.RowCount;
            columnIndexToReadTo = reader.ColumnCount;
        }

        private void UpdateCellsState(Cell cell)
        {
            currentRowIndex = reader.CurrentCellRowIndex;
            currentColumnIndex = reader.CurrentCellColumnIndex;
            currentCellValue = reader.GetCellValue(cell);
        }
    }
}
