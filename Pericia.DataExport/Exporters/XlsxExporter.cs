using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Pericia.DataExport.Exporters
{
    internal class XlsxExporter : IFormatExporter
    {
        private MemoryStream stream;
        private SpreadsheetDocument package;

        int currentCol = 0;
        int currentRow = 0;
        Row row;
        SheetData sheetData;

        public XlsxExporter()
        {
            stream = new MemoryStream();
            package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            InitWorkbook();
        }

        private void InitWorkbook()
        {
            WorkbookPart workbookPart = package.AddWorkbookPart();

            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbookPart.Workbook = workbook;

            Sheets sheets = new Sheets();
            workbook.Append(sheets);

            Sheet sheet = new Sheet() { Name = "Sheet1", SheetId = (UInt32Value)1U, Id = "rId1" };
            sheets.Append(sheet);

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            Worksheet worksheet = new Worksheet();
            sheetData = new SheetData();
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            NewLine();
        }

        public void NewLine()
        {
            currentCol = 1;
            currentRow++;
            row = new Row();
            sheetData.Append(row);
        }

        public void WriteData(string data)
        {
            Cell cell = new Cell()
            {
                CellReference = ExcelColumnFromNumber(currentCol++) + currentRow.ToString(),
                DataType = CellValues.InlineString
            };
            InlineString inlineString = new InlineString();
            Text text = new Text();
            text.Text = data;
            inlineString.Append(text);
            cell.Append(inlineString);
            row.Append(cell);
        }

        private static Dictionary<int, string> excelColumnCache = new Dictionary<int, string>();
        private static string ExcelColumnFromNumber(int column)
        {
            if (excelColumnCache.ContainsKey(column))
            {
                return excelColumnCache[column];
            }

            string columnString = "";
            decimal columnNumber = column;
            while (columnNumber > 0)
            {
                decimal currentLetterNumber = (columnNumber - 1) % 26;
                char currentLetter = (char)(currentLetterNumber + 65);
                columnString = currentLetter + columnString;
                columnNumber = (columnNumber - (currentLetterNumber + 1)) / 26;
            }

            excelColumnCache.Add(column, columnString);
            return columnString;
        }


        public Stream GetStream()
        {
            package.Dispose();

            stream.Position = 0;
            return stream;
        }

    }
}
