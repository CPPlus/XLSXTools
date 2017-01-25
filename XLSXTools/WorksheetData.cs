using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXTools
{
    public class WorksheetData
    {
        public OpenXmlWriter Writer { get; private set; }
        public WorksheetPart WorksheetPart { get; private set; }
        public int LastCellRowIndex { get; set; }
        public int LastCellColumnIndex { get; set; }

        public WorksheetData(WorkbookPart parentWorkbookPart) 
        {
            WorksheetPart = parentWorkbookPart.AddNewPart<WorksheetPart>();
            Writer = OpenXmlWriter.Create(WorksheetPart);

            LastCellRowIndex = 1;
            LastCellColumnIndex = 1;
        }
    }
}
