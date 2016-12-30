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
        private string usedRange;

        public XLSXReader(string path)
        {
            spreadsheetDocument = SpreadsheetDocument.Open(path, true);
            workbookPart = spreadsheetDocument.WorkbookPart;
            worksheetPart = workbookPart.WorksheetParts.First();

            TrackDateStyleIndices();

            openXmlReader = OpenXmlReader.Create(worksheetPart);
            // ReadUntilDimensionData();

            // GetRowCount();
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

        public int GetRowCount()
        {
            string endCellReference = usedRange.Split(':')[1];
            return XLSXUtils.CellReferenceToRowIndex(endCellReference);
        }

        public void Close()
        {
            if (openXmlReader != null)
            {
                openXmlReader.Close();
                openXmlReader = null;
            }
            
            spreadsheetDocument.Close();
        }

        private void ReadUntilDimensionData()
        {
            while (openXmlReader.Read() && openXmlReader.ElementType != typeof(SheetDimension)) { }
            SheetDimension sheetDimension = openXmlReader.LoadCurrentElement() as SheetDimension;
            usedRange = sheetDimension.Reference;
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
