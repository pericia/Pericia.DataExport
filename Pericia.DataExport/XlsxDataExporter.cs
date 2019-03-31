using DocumentFormat.OpenXml;
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

        private Lazy<uint> dateStyleIndex;

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

            dateStyleIndex = new Lazy<uint>(GenerateStylesheet);
        }

        private uint GenerateStylesheet()
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
            WorkbookStylesPart workbookStylesPart1 = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac x16r2 xr" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
            stylesheet1.AddNamespaceDeclaration("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");
            stylesheet1.AddNamespaceDeclaration("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision");

            Fonts fonts1 = new Fonts() { Count = 1U, KnownFonts = true };
            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = 1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };
            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontScheme1);
            fonts1.Append(font1);

            Fills fills1 = new Fills() { Count = 1U };
            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };
            fill1.Append(patternFill1);
            fills1.Append(fill1);

            Borders borders1 = new Borders() { Count = 0U };
            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();
            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);
            borders1.Append(border1);

            CellStyleFormats cellStyleFormats1 = new CellStyleFormats() { Count = 1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = 0U };
            cellStyleFormats1.Append(cellFormat1);

            CellFormats cellFormats1 = new CellFormats() { Count = 2U };
            CellFormat cellFormat43 = new CellFormat() { NumberFormatId = 0U, FormatId = 0U };
            CellFormat cellFormat44 = new CellFormat() { NumberFormatId = 14U, FormatId = 0U, ApplyNumberFormat = true };

            cellFormats1.Append(cellFormat43);
            cellFormats1.Append(cellFormat44);

            stylesheet1.Append(fonts1);
            stylesheet1.Append(fills1);
            stylesheet1.Append(borders1);
            stylesheet1.Append(cellStyleFormats1);
            stylesheet1.Append(cellFormats1);

            workbookStylesPart1.Stylesheet = stylesheet1;
            return 1;
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
                cell.DataType = CellValues.Date;
                cell.StyleIndex = dateStyleIndex.Value;
                cell.CellValue = new CellValue(((DateTime)data).ToString("s", CultureInfo.InvariantCulture));
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
