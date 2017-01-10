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

        public XLSXRowReader(string srcFilePath) : this(srcFilePath, "Sheet1")
        {

        }

        public XLSXRowReader(string srcFilePath, string sheetName)
        {
            this.srcFilePath = srcFilePath;
            this.sheetName = sheetName;

            ResetState();
        }

        public void Close()
        {
            reader.Close();
        }

        public bool ReadNextRecord(out string[] record)
        {
            if (!couldReadLast)
            {
                FillWithEmptyRecord(out record);
                return false;
            }

            string[] result = new string[columnIndexToReadTo];
            rowIndexToRead++;

            bool hasWrittenValue = false;
            do
            {
                UpdateCellsState(reader.CurrentCell);

                if (currentRowIndex == rowIndexToRead)
                {
                    result[currentColumnIndex - 1] = currentCellValue;
                    hasWrittenValue = true;
                }
                else
                {
                    if (!hasWrittenValue)
                        FillWithEmptyRecord(out result);
                    hasWrittenValue = true;
                    break;
                }
            } while (couldReadLast = reader.ReadNextCell());

            if (!couldReadLast)
            {
                if (!hasWrittenValue)
                {
                    FillWithEmptyRecord(out record);
                    return false;
                }
            }
            
            record = result;

            return true;
        }

        private void FillWithEmptyRecord(out string[] record)
        {
            record = new string[reader.ColumnCount];
            for (int i = 0; i < record.Length; i++)
            {
                record[i] = string.Empty;
            }
        }

        private void ResetState()
        {
            reader = new XLSXReader(srcFilePath, sheetName);
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
