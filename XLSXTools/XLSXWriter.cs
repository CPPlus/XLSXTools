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
    public class XLSXWriter
    {
        private SpreadsheetDocument spreadsheetDocument;
        private WorkbookPart workbookPart;
        private WorksheetPart worksheetPart;
        private SharedStringTablePart sharedStringTablePart;
        private uint lastSheetId = 1;
        private int lastCellRowIndex = 1;
        private int lastCellColumnIndex = 1;
        private OpenXmlWriter openXmlWriter;
        public int RowsWritten { get; private set; }

        public XLSXWriter(string path)
        {
            spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
            workbookPart = spreadsheetDocument.AddWorkbookPart();
            worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

            WriteWorkbookPart();
            WriteSharedStringTablePart();
            WriteWorkbookStylesPart();

            RowsWritten = 0;
        }

        // Call before writing anything.
        public void Start()
        {
            openXmlWriter = OpenXmlWriter.Create(worksheetPart);

            openXmlWriter.WriteStartElement(new Worksheet());
            openXmlWriter.WriteStartElement(new SheetData());
            openXmlWriter.WriteStartElement(new Row());
        }

        public void JumpForwardTo(string cellReference)
        {
            int rowIndex = XLSXUtils.CellReferenceToRowIndex(cellReference);
            for (int i = lastCellRowIndex; i < rowIndex; i++)
            {
                NewRow();
            }

            int columnIndex = XLSXUtils.CellReferenceToColumnIndex(cellReference);
            lastCellColumnIndex = columnIndex;
        }

        // Writes a value to the cell to the right of 
        // the last cell we have written to.
        public string Write(int value)
        {
            return Write(value, 0, false);
        }

        public string Write(int value, uint styleIndex)
        {
            return Write(value, styleIndex, false);
        }

        public string Write(string value)
        {
            int sharedStringIndex = InsertSharedStringItem(value);
            return Write(sharedStringIndex, 0, true);
        }

        public string Write(string value, uint styleIndex)
        {
            int sharedStringIndex = InsertSharedStringItem(value);
            return Write(sharedStringIndex, styleIndex, true);
        }

        public string Write(int value, uint styleIndex = 0, bool isSharedString = false)
        {
            return Write((decimal)value, styleIndex, isSharedString);
        }

        public string Write(decimal value, uint styleIndex = 0, bool isSharedString = false)
        {
            Cell cell = new Cell();

            List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>();
            string cellReference = GetCurrentCellReference();
            attributes.Add(new OpenXmlAttribute("r", null, cellReference));
            if (styleIndex != 0)
            {
                attributes.Add(new OpenXmlAttribute("s", null, styleIndex.ToString()));
            }
            if (isSharedString)
            {
                attributes.Add(new OpenXmlAttribute("t", null, "s"));
            }

            CellValue cellValue = new CellValue(value.ToString());
            cell.Append(cellValue);

            openXmlWriter.WriteStartElement(cell, attributes);
            openXmlWriter.WriteElement(cellValue);
            openXmlWriter.WriteEndElement();

            return cellReference;
        }

        public string WriteInline(string value, uint styleIndex = 0)
        {
            Cell cell = new Cell();

            List<OpenXmlAttribute> attributes = new List<OpenXmlAttribute>();
            string cellReference = GetCurrentCellReference();
            attributes.Add(new OpenXmlAttribute("r", null, cellReference));
            attributes.Add(new OpenXmlAttribute("t", null, "inlineStr"));
            if (styleIndex != 0)
            {
                attributes.Add(new OpenXmlAttribute("s", null, styleIndex.ToString()));
            }

            Text cellValue = new Text(value);

            InlineString inlineString = new InlineString();

            openXmlWriter.WriteStartElement(cell, attributes);
            openXmlWriter.WriteStartElement(inlineString);
            openXmlWriter.WriteElement(cellValue);
            openXmlWriter.WriteEndElement();
            openXmlWriter.WriteEndElement();

            return cellReference;
        }

        public void NewRow()
        {
            openXmlWriter.WriteEndElement();
            openXmlWriter.WriteStartElement(new Row());
            lastCellRowIndex++;
            lastCellColumnIndex = 1;

            RowsWritten++;
        }

        // Call after last call to 'Write'.
        public void Finish()
        {
            openXmlWriter.WriteEndElement();
            openXmlWriter.WriteEndElement();
            openXmlWriter.WriteEndElement();
        }

        public void Close()
        {
            if (openXmlWriter != null)
                openXmlWriter.Close();

            if (spreadsheetDocument != null)
                spreadsheetDocument.Close();
        }

        private void WriteWorkbookPart()
        {
            OpenXmlWriter openXmlWriter = OpenXmlWriter.Create(spreadsheetDocument.WorkbookPart);

            openXmlWriter.WriteStartElement(new Workbook());
            openXmlWriter.WriteStartElement(new Sheets());

            AddSheets(openXmlWriter);
            
            openXmlWriter.WriteEndElement();
            openXmlWriter.WriteEndElement();

            openXmlWriter.Close();
        }

        private void AddSheets(OpenXmlWriter openXmlWriter)
        {
            AddSheet("Sheet1", openXmlWriter);
        }

        private string GetCurrentCellReference()
        {
            return XLSXUtils.ColumnIndexToLetter(lastCellColumnIndex++) + lastCellRowIndex;
        }

        private void AddSheet(string name, OpenXmlWriter openXmlWriter)
        {
            Sheet sheet = new Sheet()
            {
                Name = name,
                SheetId = lastSheetId++,
                Id = workbookPart.GetIdOfPart(worksheetPart)
            };
            openXmlWriter.WriteElement(sheet);
        }

        private void WriteSharedStringTablePart()
        {
            sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
        }

        private int InsertSharedStringItem(string text)
        {
            int i = 0;
            foreach (SharedStringItem item in sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }
            
            sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            sharedStringTablePart.SharedStringTable.Save();

            return i;
        }

        private void WriteWorkbookStylesPart()
        {
            WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = new Stylesheet();
            workbookStylesPart.Stylesheet = stylesheet;

            // Fonts.
            Font defaultFont = new Font();

            Font boldFont = new Font();        
            Bold bold = new Bold();
            boldFont.Append(bold);

            Fonts fonts = new Fonts();
            fonts.Append(defaultFont);
            fonts.Append(boldFont);
            fonts.Count = 2;

            // Fills.
            Fill defaultFill = new Fill()
            {
                PatternFill = new PatternFill { PatternType = PatternValues.None }
            };
            
            Fill defaultFill2 = new Fill()
            {
                PatternFill = new PatternFill { PatternType = PatternValues.Gray125 }
            };

            PatternFill yellowPatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid 
            };
            yellowPatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFFFF00") }; // red fill
            yellowPatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            Fill yellowFill = new Fill()
            {
                PatternFill = yellowPatternFill
            };

            PatternFill bluePatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid
            };
            bluePatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("E5E5FF") }; // blue fill
            bluePatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            Fill blueFill = new Fill()
            {
                PatternFill = bluePatternFill
            };

            PatternFill redPatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid
            };
            redPatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("FF5555") }; // red fill
            redPatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            Fill redFill = new Fill()
            {
                PatternFill = redPatternFill
            };

            PatternFill greenPatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid
            };
            greenPatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("C6E0B4") }; // green fill
            greenPatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            Fill greenFill = new Fill()
            {
                PatternFill = greenPatternFill
            };

            Fills fills = new Fills();             
            fills.Append(defaultFill);
            fills.Append(defaultFill2);
            fills.Append(yellowFill);
            fills.Append(blueFill);
            fills.Append(redFill);
            fills.Append(greenFill);
            fills.Count = 6;

            // Borders.
            Border defaultBorder = new Border();   

            Borders borders = new Borders();
            borders.Append(defaultBorder);
            borders.Count = 1;

            // Cell formats
            CellFormat defaultCellFormat = new CellFormat();
            CellFormat yellowFillCellFormat = new CellFormat() { FontId = 0, FillId = 2, BorderId = 0, ApplyFill = true };
            CellFormat blueFillCellFormat = new CellFormat() { FontId = 0, FillId = 3, BorderId = 0, ApplyFill = true };
            CellFormat redFillCellFormat = new CellFormat() { FontId = 0, FillId = 4, BorderId = 0, ApplyFill = true };
            CellFormat greenFillCellFormat = new CellFormat() { FontId = 0, FillId = 5, BorderId = 0, ApplyFill = true };

            CellFormats cellFormats = new CellFormats();
            cellFormats.Append(defaultCellFormat);
            cellFormats.Append(yellowFillCellFormat);
            cellFormats.Append(blueFillCellFormat);
            cellFormats.Append(redFillCellFormat);
            cellFormats.Append(greenFillCellFormat);
            cellFormats.Count = 5;

            // Append fonts, fills, borders and cell formats.
            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellFormats);
            
            // Save.
            workbookStylesPart.Stylesheet.Save();
        }


    }
}
