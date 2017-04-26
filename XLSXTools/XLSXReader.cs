using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
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
        private Dictionary<uint, string> customNumberFormatIds;
        public string UsedRange { get; private set; }

        private int rowCount = -1;
        private int columnCount = -1;

        private string path;
        private string sheetName;

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

        public XLSXReader(string path, bool customRangeCalculation = false, int rowStart = 0) : this(path, "Sheet1", customRangeCalculation, rowStart)
        {

        }

        public XLSXReader(string path, string sheet, bool customRangeCalculation = false, int rowStart = 0)
        {
            this.path = path;
            this.sheetName = sheet;

            SetupReader();

            // Either calculate the true used range yourself or trust the excel one.
            if (customRangeCalculation)
            {
                CalculateUsedRange(rowStart);
                SetupReader();
            }
            else ReadUntilDimensionData();
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

        public string GetCellFormat(Cell cell)
        {
            bool isFormatted = (cell.StyleIndex != null) && styleIndicesWithNumberFormat.ContainsKey(cell.StyleIndex);
            if (isFormatted)
            {
                string formatString = null;
                uint numberFormatId;
                styleIndicesWithNumberFormat.TryGetValue(cell.StyleIndex, out numberFormatId);

                switch (numberFormatId)
                {
                    case 0: formatString = "General"; break;
                    case 1: formatString = "0"; break;
                    case 2: formatString = "0.00"; break;
                    case 3: formatString = "#,##0"; break;
                    case 4: formatString = "#,##0.00"; break;
                    case 9: formatString = "0%"; break;
                    case 10: formatString = "0.00%"; break;
                    case 11: formatString = "0.00E+00"; break;
                    case 12: formatString = "# ?/?"; break;
                    case 13: formatString = "# ??/??"; break;
                    case 14: formatString = "d/m/yyyy"; break;
                    case 15: formatString = "d-mmm-yy"; break;
                    case 16: formatString = "d-mmm"; break;
                    case 17: formatString = "mmm-yy"; break;
                    case 18: formatString = "h:mm tt"; break;
                    case 19: formatString = "h:mm:ss tt"; break;
                    case 20: formatString = "H:mm"; break;
                    case 21: formatString = "H:mm:ss"; break;
                    case 22: formatString = "m/d/yyyy H:mm"; break;
                    case 37: formatString = "#,##0;(#,##0)"; break;
                    case 38: formatString = "#,##0;[Red](#,##0)"; break;
                    case 39: formatString = "#,##0.00;(#,##0.00)"; break;
                    case 40: formatString = "#,##0.00;[Red](#,##0.00)"; break;
                    case 45: formatString = "mm:ss"; break;
                    case 46: formatString = "[h]:mm:ss"; break;
                    case 47: formatString = "mmss.0"; break;
                    case 48: formatString = "##0.0E+0"; break;
                    case 49: formatString = "@"; break;
                    default:
                        customNumberFormatIds.TryGetValue(numberFormatId, out formatString);
                        break;
                }
                return formatString;
            }
            else return null;
        }
       
        public string GetCellValue(Cell cell)
        {
            string value = string.Empty;

            if (cell != null)
            {
                value = cell.InnerText;

                string formatString = GetCellFormat(cell);
                
                // If the format is not null and it is not a shared string.
                if (formatString != null && (cell.DataType == null ? true : cell.DataType.Value != CellValues.SharedString))
                {
                    double result;
                    bool success = double.TryParse(value, out result);
                    if (success)
                    {
                        if (IsValidDateTimeFormat(formatString))
                        {
                            string formatOfOutputDate = "M/d/yyyy";
                            // string formatOfOutputDate = formatString.Replace('m', 'M');

                            value = DateTime.FromOADate(result).ToString(formatOfOutputDate);
                            // value = FromExcelSerialDate(result).ToString(formatString);
                        }
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

        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
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

        public bool IsValidDateTimeFormat(string format)
        {
            return format.IndexOf('m') >= 0 && format.IndexOf('d') >= 0 && format.IndexOf('y') >= 0;
            /*
            string strDummy = DateTime.Now.ToString(format);
            DateTime dateDummy;
            bool success = DateTime.TryParseExact(strDummy, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateDummy);
            return success;
            */
        }

        private void CalculateUsedRange(int rowStart)
        {
            int rowCount = 0;
            int columnCount = 0;

            while (openXmlReader.Read())
            {
                if (openXmlReader.ElementType == typeof(Cell))
                {
                    Cell cell = (Cell)openXmlReader.LoadCurrentElement();
                    string cellValue = GetCellValue(cell);
                    
                    if (!cellValue.Equals(string.Empty))
                    {
                        string cellAddress = GetCellReference(cell);
                        int rowIndex = XLSXUtils.CellReferenceToRowIndex(cellAddress);

                        if (rowIndex >= rowStart)
                        {
                            int columnIndex = XLSXUtils.CellReferenceToColumnIndex(cellAddress);

                            if (rowIndex > rowCount) rowCount = rowIndex;
                            if (columnIndex > columnCount) columnCount = columnIndex;
                        }
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
            worksheetPart = FindSheet(sheetName);

            TrackDateStyleIndices();
            TrackCustomNumberFormatIds();

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

        private void TrackCustomNumberFormatIds()
        {
            customNumberFormatIds = new Dictionary<uint, string>();

            WorkbookStylesPart workbookStylesPart = workbookPart.GetPartsOfType<WorkbookStylesPart>().FirstOrDefault();
            if (workbookStylesPart != null)
            {
                Stylesheet stylesheet = workbookStylesPart.Stylesheet;
                NumberingFormats numberingFormats = stylesheet.NumberingFormats;
                if (numberingFormats != null)
                {
                    foreach (NumberingFormat numberingFormat in numberingFormats)
                    {
                        customNumberFormatIds.Add(numberingFormat.NumberFormatId, numberingFormat.FormatCode);
                    }
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
