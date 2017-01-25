using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXTools
{
    public class XLSXReader
    {
        private SpreadsheetDocument spreadsheetDocument;
        private WorkbookPart workbookPart;
        private WorksheetPart worksheetPart;

        private OpenXmlReader openXmlReader;
        public Cell CurrentCell { get; private set; }

        private Dictionary<uint, uint> styleIndicesWithNumberFormat;
        public string UsedRange { get; private set; }

        private int rowCount = -1;
        private int columnCount = -1;

        private string path;

        public string CurrentCellAddress
        {
            get
            {
                return GetCellReference(CurrentCell);
            }
        }

        public int CurrentCellColumnIndex
        {
            get
            {
                return XLSXUtils.CellReferenceToColumnIndex(CurrentCellAddress);
            }
        }

        public int CurrentCellRowIndex
        {
            get
            {
                return XLSXUtils.CellReferenceToRowIndex(CurrentCellAddress);
            }
        }

        public int RowCount
        {
            get
            {
                if (rowCount == -1)
                {
                    string endCellReference = UsedRange.Split(':')[1];
                    rowCount = XLSXUtils.CellReferenceToRowIndex(endCellReference);
                }
                return rowCount;
            }
        }

        public int ColumnCount
        {
            get
            {
                if (columnCount == -1)
                {
                    string endCellReference = UsedRange.Split(':')[1];
                    columnCount = XLSXUtils.CellReferenceToColumnIndex(endCellReference);
                }
                return columnCount;
            }
        }

        public bool EOF
        {
            get
            {
                return openXmlReader.EOF;
            }
        }

        public XLSXReader(string path) : this(path, "Sheet1")
        {

        }

        public XLSXReader(string path, string sheet)
        {
            this.path = path;
            
            SetupReader();
            SetSheet(sheet);
            //CalculateUsedRange();
            ReadUntilDimensionData();

            SetupReader();
        }

        public void SetSheet(string sheetName)
        {
            WorksheetPart worksheetPart = FindSheet(sheetName);
            if (worksheetPart != null)
            {
                this.worksheetPart = worksheetPart;

                if (openXmlReader != null)
                {
                    openXmlReader.Close();
                    openXmlReader = null;
                }

                openXmlReader = OpenXmlReader.Create(this.worksheetPart);
            }
            else
            {
                Console.WriteLine("Worksheet doesn't exist.");
            }
        }

        public bool ReadNextCell()
        {
            bool isDataToRead = false;
            while ((isDataToRead = openXmlReader.Read()) && openXmlReader.ElementType != typeof(Cell)) {}
            if (!openXmlReader.EOF)
            {
                CurrentCell = openXmlReader.LoadCurrentElement() as Cell;
            } else
            {
                CurrentCell = null;
            }
            
            return isDataToRead;
        }
       
        public string GetCellValue(Cell cell)
        {
            string value = string.Empty;

            if (cell != null)
            {
                value = cell.InnerText;

                // If it's a formatted integer.
                bool isFormatted = (cell.StyleIndex != null) && styleIndicesWithNumberFormat.ContainsKey(cell.StyleIndex);
                if (isFormatted)
                {
                    uint numberFormatId;
                    bool success = styleIndicesWithNumberFormat.TryGetValue(cell.StyleIndex, out numberFormatId);
                    if (success)
                    {
                        string formatString = null;

                        switch (numberFormatId)
                        {
                            case 14: formatString = "M/d/yyyy"; break;
                            case 15: formatString = "d-mmm-yy"; break;
                            case 16: formatString = "d-mmm"; break;
                            case 17: formatString = "mmm-yy"; break;
                            case 18: formatString = "h:mm tt"; break;
                            case 19: formatString = "h:mm:ss tt"; break;
                            case 20: formatString = "H:mm"; break;
                            case 21: formatString = "H:mm: ss"; break;
                        }

                        if (formatString != null)
                        {
                            int result;
                            success = int.TryParse(value, out result);
                            if (success)
                            {
                                value = DateTime.FromOADate(result).ToString(formatString);
                            }
                        }
                            
                    } else
                    {
                        Console.WriteLine("Key not found.");
                    }
                }

                // If it's a string or boolean.
                if (cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                            if (stringTable != null)
                            {
                                int result;
                                bool success = int.TryParse(value, out result);
                                if (success)
                                    value = stringTable.SharedStringTable.ElementAt(result).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }

        public string GetCellReference(Cell cell)
        {
            return cell.CellReference.Value;
        }

        public void Close()
        {
            if (openXmlReader != null)
            {
                openXmlReader.Close();
                openXmlReader = null;
            }
            
            if (spreadsheetDocument != null)
            {
                spreadsheetDocument.Close();
                spreadsheetDocument = null;
            }
        }

        private void CalculateUsedRange()
        {
            int rowCount = 0;
            int columnCount = 0;

            while (openXmlReader.Read())
            {
                if (openXmlReader.ElementType == typeof(Cell))
                {
                    Cell cell = (Cell)openXmlReader.LoadCurrentElement();
                    string cellValue = GetCellValue(cell);
                    
                    if (!cellValue.Equals(""))
                    {
                        string cellAddress = GetCellReference(cell);
                        int rowIndex = XLSXUtils.CellReferenceToRowIndex(cellAddress);

                        int columnIndex = XLSXUtils.CellReferenceToColumnIndex(cellAddress);
                        if (rowIndex > rowCount) rowCount = rowIndex;
                        if (columnIndex > columnCount) columnCount = columnIndex;
                    }
                }
            }

            string columnCountLetter = XLSXUtils.ColumnIndexToLetter(columnCount);
            UsedRange = string.Format("A1:{0}{1}", columnCountLetter, rowCount);
        }

        private void SetupReader()
        {
            Close();

            spreadsheetDocument = SpreadsheetDocument.Open(path, true);
            workbookPart = spreadsheetDocument.WorkbookPart;
            worksheetPart = workbookPart.WorksheetParts.First();

            TrackDateStyleIndices();

            openXmlReader = OpenXmlReader.Create(worksheetPart);
        }

        private void ReadUntilDimensionData()
        {
            while (openXmlReader.Read() && openXmlReader.ElementType != typeof(SheetDimension)) { }
            SheetDimension sheetDimension = openXmlReader.LoadCurrentElement() as SheetDimension;
            UsedRange = sheetDimension.Reference;
        }

        private void TrackDateStyleIndices()
        {
            styleIndicesWithNumberFormat = new Dictionary<uint, uint>();

            WorkbookStylesPart workbookStylesPart = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (workbookStylesPart != null)
            {
                Stylesheet stylesheet = workbookStylesPart.Stylesheet;
                CellFormats cellFormats = stylesheet.CellFormats;
                uint cellFormatIndex = 0;
                foreach (CellFormat cellFormat in cellFormats)
                {
                    BooleanValue applyNumberFormat = cellFormat.ApplyNumberFormat;
                    if (applyNumberFormat != null && applyNumberFormat.Value)
                    {
                        styleIndicesWithNumberFormat.Add(cellFormatIndex, cellFormat.NumberFormatId);
                    }
                    cellFormatIndex++;
                }
            }
        }

        private WorksheetPart FindSheet(string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
            if (sheet == null) return null;
            else return workbookPart.GetPartById(sheet.Id) as WorksheetPart;
        }
    }
}
