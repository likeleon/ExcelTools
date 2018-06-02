using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace ExcelTools
{
    public sealed partial class Program
    {
        private const string SampleExcelFile = "SampleData.xlsx";

        public static void Main(string[] args)
        {
            using (var sheetDocument = SpreadsheetDocument.Open(SampleExcelFile, isEditable: false))
            {
                var workbookPart = sheetDocument.WorkbookPart;
                var lazyNumberingFormatCodeById = GetLazyNumberingFormatCodeById(workbookPart.WorkbookStylesPart.Stylesheet);

                foreach (var sheet in workbookPart.Workbook.Sheets.OfType<Sheet>())
                {
                    var worksheetPart = workbookPart.GetPartById(sheet.Id.Value) as WorksheetPart;
                    if (worksheetPart == null)
                    {
                        throw new Exception($"WorksheetPart not found with id '{sheet.Id}'. Sheet name is '{sheet.Name}'.");
                    }

                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                    if (sheetData == null)
                    {
                        continue;
                    }

                    foreach (var table in worksheetPart.TableDefinitionParts.Select(p => p.Table))
                    {
                        var dataTable = ReadSpreadsheetTable(table, sheetData, workbookPart, lazyNumberingFormatCodeById);
                    }
                }
            }
        }

        private static Lazy<IReadOnlyDictionary<uint, string>> GetLazyNumberingFormatCodeById(Stylesheet stylesheet)
        {
            return new Lazy<IReadOnlyDictionary<uint, string>>(() =>
            {
                var dictionary = new Dictionary<uint, string>();
                foreach (var customFormat in stylesheet.NumberingFormats.OfType<NumberingFormat>())
                {
                    dictionary.Add(customFormat.NumberFormatId.Value, customFormat.FormatCode);
                };
                foreach (var kvp in PredefinedNumberingFormats.FormatCodes)
                {
                    dictionary.Add((uint)kvp.Key, kvp.Value);
                }
                return dictionary;
            });
        }

        private static DataTable ReadSpreadsheetTable(Table table, SheetData sheetData, WorkbookPart workbookPart, Lazy<IReadOnlyDictionary<uint, string>> lazyNumberingFormatCodeById)
        {
            var dataTable = new DataTable(table.DisplayName);

            foreach (var tableColumn in table.TableColumns.OfType<TableColumn>())
            {
                dataTable.Columns.Add(tableColumn.Name.Value);
            }

            var tableReference = TableReference.Parse(table.Reference);

            int startRowIndex = tableReference.StartCell.RowIndex + 1; // "+1": 열 헤더 제외
            int endRowIndex = tableReference.EndCell.RowIndex;
            int startColumnIndex = tableReference.StartCell.ColumnIndex;
            int endColumnIndex = tableReference.EndCell.ColumnIndex;

            foreach (var sheetRow in sheetData.Elements<Row>())
            {
                int rowIndex = int.Parse(sheetRow.RowIndex);
                if (rowIndex < startRowIndex || rowIndex > endRowIndex)
                {
                    continue;
                }

                var row = dataTable.NewRow();

                foreach (var sheetCell in sheetRow.Elements<Cell>())
                {
                    var cellReference = CellReference.Parse(sheetCell.CellReference);
                    if (cellReference.ColumnIndex < startColumnIndex || cellReference.ColumnIndex > endColumnIndex)
                    {
                        continue;
                    }

                    row[cellReference.ColumnIndex - 1] = GetCellValue(sheetCell, workbookPart, lazyNumberingFormatCodeById);
                }

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart, Lazy<IReadOnlyDictionary<uint, string>> lazyNumberingFormatsById)
        {
            string rawText = cell.CellValue.InnerXml;

            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var sharedStringTable = workbookPart.SharedStringTablePart.SharedStringTable;
                var index = int.Parse(rawText);
                return sharedStringTable.ChildElements[index].InnerText;
            }
            else if (cell.StyleIndex != null)
            {
                var stylesheet = workbookPart.WorkbookStylesPart.Stylesheet;
                var cellFormat = (CellFormat)stylesheet.CellFormats.ChildElements[(int)cell.StyleIndex.Value];
                var numberFormatId = cellFormat.NumberFormatId.Value;
                var numberingFormatCode = lazyNumberingFormatsById.Value[numberFormatId];
                return GetFormattedText(rawText, numberFormatId, numberingFormatCode);
            }
            else
            {
                return rawText;
            }
        }

        /// <summary>
        /// <see cref="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx"/>>
        /// </summary>
        private static string GetFormattedText(string rawText, uint formatId, string formatCode)
        {
            double doubleValue;
            if (!double.TryParse(rawText, out doubleValue))
            {
                return rawText;
            }

            var actualFormat = GetActualFormat(formatCode, doubleValue);

            if (IsDateTimeFormat(formatId))
            {
                return DateTime.FromOADate(doubleValue).ToString(actualFormat);
            }
            else
            {
                return doubleValue.ToString(actualFormat);
            }
        }

        private static string GetActualFormat(string formatCode, double value)
        {
            // The format is actually 4 formats split by a semi-colon
            // 0 for positive, 1 for negative, 2 for zero (I'm ignoring the 4th format which is for text)
            string[] formatComponents = formatCode.Split(';');

            int elementToUse = value > 0 ? 0 : (value < 0 ? 1 : 2);

            string actualFormat = formatComponents[elementToUse];

            actualFormat = RemoveUnwantedCharacters(actualFormat, '_');
            actualFormat = RemoveUnwantedCharacters(actualFormat, '*');

            // Backslashes are an escape character it seems - I'm ignoring them
            return actualFormat.Replace("\"", ""); ;
        }

        private static string RemoveUnwantedCharacters(string excelFormat, char character)
        {
            /*  The _ and * characters are used to control lining up of characters
                they are followed by the character being manipulated so I'm ignoring
                both the _ and * and the character immediately following them.
                Note that this is buggy as I don't check for the preceeding
                backslash escape character which I probably should
                */
            int index = excelFormat.IndexOf(character);
            int occurance = 0;
            while (index != -1)
            {
                // Replace the occurance at index using substring
                excelFormat = excelFormat.Substring(0, index) + excelFormat.Substring(index + 2);
                occurance++;
                index = excelFormat.IndexOf(character, index);
            }
            return excelFormat;
        }

        private static bool IsDateTimeFormat(uint formatId)
        {
            return (14 <= formatId && formatId <= 22) || (164 <= formatId);
        }
    }
}
