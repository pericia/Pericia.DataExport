﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace Pericia.DataExport
{
    public class XlsxDataExporter : DataExporter
    {
        private SpreadsheetDocument package;

        int currentCol = 0;
        int currentRow = 0;
        Row row;
        WorkbookPart workbookPart;
        SheetData sheetData;
        Sheets sheets;
        uint sheetCount = 0;

        public XlsxDataExporter()
        {
            package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            InitWorkbook();
        }

        private void InitWorkbook()
        {
            workbookPart = package.AddWorkbookPart();

            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            workbookPart.Workbook = workbook;

            sheets = new Sheets();
            workbook.Append(sheets);

        }

        protected override void NewSheet(string name)
        {
            var sheetId = "rId" + (++sheetCount);
            name = NewSheetName(name);

            Sheet sheet = new Sheet() { Name = name, SheetId = sheetCount, Id = sheetId };
            sheets.Append(sheet);

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetId);
            Worksheet worksheet = new Worksheet();
            sheetData = new SheetData();
            worksheet.Append(sheetData);
            worksheetPart.Worksheet = worksheet;

            currentCol = 0;
            currentRow = 0;

            NewLine();
        }
        
        protected override void NewLine()
        {
            currentCol = 1;
            currentRow++;
            row = new Row();
            sheetData.Append(row);
        }

        protected override void WriteData(object data)
        {
            Cell cell = new Cell()
            {
                CellReference = ExcelColumnFromNumber(currentCol++) + currentRow.ToString(CultureInfo.InvariantCulture),
            };

            if (data is sbyte || data is byte || data is short || data is ushort || data is int || data is uint
                || data is long || data is ulong || data is float || data is double || data is decimal)
            {
                cell.DataType = CellValues.Number;
                var toStringMethod = data.GetType().GetMethod("ToString", new Type[] { typeof(IFormatProvider) });
                var textValue = (string)toStringMethod.Invoke(data, new object[] { CultureInfo.InvariantCulture });
                cell.CellValue = new CellValue(textValue);
            }
            else if (data is DateTime)
            {
                // TODO date style
                cell.DataType = CellValues.String;
                //cell.StyleIndex = 1U;
                var dateTime = (DateTime)data;
                //textValue = dateTime.ToString("s", CultureInfo.InvariantCulture);
                cell.CellValue = new CellValue(dateTime.ToShortDateString());
            }
            else if (data is bool)
            {
                cell.DataType = CellValues.Boolean;
                cell.CellValue = new CellValue((bool)data ? "1" : "0");
            }
            else
            {
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue(data.ToString());
            }

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


        public override MemoryStream GetFile()
        {
            package.Dispose();
            stream.Position = 0;
            return stream;
        }


        
        private List<string> sheetNames = new List<string>();
        protected string NewSheetName(string suggestedName)
        {
            if (suggestedName == null)
            {
                suggestedName = "Sheet" + sheetCount;
            }

            while (sheetNames.Contains(suggestedName))
            {
                Regex numberRegex = new Regex("([0-9]+)$");
                var match = numberRegex.Match(suggestedName);
                if (match.Success)
                {
                    var newCount = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture) + 1;
                    suggestedName = numberRegex.Replace(suggestedName, newCount.ToString(CultureInfo.InvariantCulture));
                }
                else
                {
                    suggestedName = suggestedName + "1";
                }
            }

            sheetNames.Add(suggestedName);
            return suggestedName;
        }

    }
}
