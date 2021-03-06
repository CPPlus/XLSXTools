﻿using DocumentFormat.OpenXml;
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
        private SharedStringTablePart sharedStringTablePart;

        private WorkbookPart workbookPart;
        private OpenXmlWriter workbookWriter;

        private Dictionary<string, WorksheetData> worksheetDataMaps;
        private WorksheetData activeWorksheetData;
        private List<WorksheetData> worksheetDatas;
        private uint lastSheetId = 0;

        public XLSXWriter(string path)
        {
            spreadsheetDocument = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);

            workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookWriter = OpenXmlWriter.Create(workbookPart);

            worksheetDataMaps = new Dictionary<string, WorksheetData>();
            worksheetDatas = new List<WorksheetData>();
            StartWorkbookPart();

            WriteSharedStringTablePart();
            WriteWorkbookStylesPart();
        }

        public void SetWorksheet(string name)
        {
            WorksheetData data = null;
            if (worksheetDataMaps.ContainsKey(name))
            {
                worksheetDataMaps.TryGetValue(name, out data);
            } else
            {
                data = new WorksheetData(workbookPart);
                worksheetDatas.Add(data);
                AddSheet(name, workbookWriter);
                worksheetDataMaps.Add(name, data);

                data.Writer.WriteStartElement(new Worksheet());
                data.Writer.WriteStartElement(new SheetData());
                data.Writer.WriteStartElement(new Row());
            }
            activeWorksheetData = data;
        }

        public void JumpForwardTo(string cellReference)
        {
            int rowIndex = XLSXUtils.CellReferenceToRowIndex(cellReference);
            for (int i = activeWorksheetData.LastCellRowIndex; i < rowIndex; i++)
            {
                NewRow();
            }

            int columnIndex = XLSXUtils.CellReferenceToColumnIndex(cellReference);
            activeWorksheetData.LastCellColumnIndex = columnIndex;
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

            activeWorksheetData.Writer.WriteStartElement(cell, attributes);
            activeWorksheetData.Writer.WriteElement(cellValue);
            activeWorksheetData.Writer.WriteEndElement();

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

            activeWorksheetData.Writer.WriteStartElement(cell, attributes);
            activeWorksheetData.Writer.WriteStartElement(inlineString);
            activeWorksheetData.Writer.WriteElement(cellValue);
            activeWorksheetData.Writer.WriteEndElement();
            activeWorksheetData.Writer.WriteEndElement();

            return cellReference;
        }

        public void NewRow()
        {
            activeWorksheetData.Writer.WriteEndElement();
            activeWorksheetData.Writer.WriteStartElement(new Row());
            activeWorksheetData.LastCellRowIndex++;
            activeWorksheetData.LastCellColumnIndex = 1;
        }

        // Call after last call to 'Write'.
        public void Finish()
        {
            for (int i = 0; i < worksheetDatas.Count; i++)
            {
                worksheetDatas[i].Writer.WriteEndElement();
                worksheetDatas[i].Writer.WriteEndElement();
                worksheetDatas[i].Writer.WriteEndElement();
            }

            FinishWorkbookPart();
        }

        public void Close()
        {
            for (int i = 0; i < worksheetDatas.Count; i++)
                worksheetDatas[i].Writer.Close();

            if (spreadsheetDocument != null)
                spreadsheetDocument.Close();
        }

        private void StartWorkbookPart()
        {
            workbookWriter.WriteStartElement(new Workbook());
            workbookWriter.WriteStartElement(new Sheets());
        }

        private void FinishWorkbookPart()
        {
            workbookWriter.WriteEndElement();
            workbookWriter.WriteEndElement();

            workbookWriter.Close();
        }

        private string GetCurrentCellReference()
        {
            return XLSXUtils.ColumnIndexToLetter(activeWorksheetData.LastCellColumnIndex++) + activeWorksheetData.LastCellRowIndex;
        }

        private void AddSheet(string name, OpenXmlWriter openXmlWriter)
        {
            Sheet sheet = new Sheet()
            {
                Name = name,
                SheetId = lastSheetId + 1,
                Id = workbookPart.GetIdOfPart(worksheetDatas[(int)lastSheetId].WorksheetPart)
            };
            lastSheetId++;
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

            PatternFill grayPatternFill = new PatternFill()
            {
                PatternType = PatternValues.Solid
            };
            grayPatternFill.ForegroundColor = new ForegroundColor { Rgb = HexBinaryValue.FromString("d7d7d7") }; // gray fill
            grayPatternFill.BackgroundColor = new BackgroundColor { Indexed = 64 };
            Fill grayFill = new Fill()
            {
                PatternFill = grayPatternFill
            };

            Fills fills = new Fills();             
            fills.Append(defaultFill);
            fills.Append(defaultFill2);
            fills.Append(yellowFill);
            fills.Append(blueFill);
            fills.Append(redFill);
            fills.Append(greenFill);
            fills.Append(grayFill);
            fills.Count = 7;

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
            CellFormat grayFillCellFormat = new CellFormat() { FontId = 0, FillId = 6, BorderId = 0, ApplyFill = true };

            CellFormats cellFormats = new CellFormats();
            cellFormats.Append(defaultCellFormat);
            cellFormats.Append(yellowFillCellFormat);
            cellFormats.Append(blueFillCellFormat);
            cellFormats.Append(redFillCellFormat);
            cellFormats.Append(greenFillCellFormat);
            cellFormats.Append(grayFillCellFormat);
            cellFormats.Count = 6;

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
