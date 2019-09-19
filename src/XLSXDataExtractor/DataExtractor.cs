using ClosedXML.Excel;
using ClosedXML.Excel.CalcEngine;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using XLSXDataExtractor.Common;
using XLSXDataExtractor.Models;

namespace XLSXDataExtractor
{
    public class DataExtractor
    {
        public XLWorkbook RequiredWorkbook { get; private set; }
        private string requiredWorkbookPath;
        public DataExtractor(string workbookPath)
        {
            requiredWorkbookPath = workbookPath;
            try
            {
                RequiredWorkbook = new XLWorkbook(workbookPath);
            }
            catch (ArgumentException) //should catch if given file is not a valid spreadsheet that can be opened
            {
                throw;
            }
        }

        public IEnumerable<IEnumerable<KeyValuePair<string, fieldValueType>>> RetrieveDataCollectionFromAllWorksheets<fieldValueType>(IEnumerable<ExtractionRequest> extractionRequests)
        {
            foreach (var worksheet in RequiredWorkbook.Worksheets)
            {
                var extractedDatas = new List<KeyValuePair<string, fieldValueType>>();
                foreach (var extractionRequest in extractionRequests)
                {
                    var extractedData = RetrieveDataFromWorkbook<fieldValueType>(worksheet.Name, extractionRequest.FieldName, extractionRequest.RowNum, extractionRequest.ColNum);
                    extractedDatas.Add(extractedData);
                }
                yield return extractedDatas;
            }
        }

        public IEnumerable<KeyValuePair<string, fieldValueType>> RetrieveDataCollectionFromSpecificWorksheet<fieldValueType>(string workSheetName, IEnumerable<ExtractionRequest> extractionRequests)
        {
            foreach (var extractionRequest in extractionRequests)
            {
                var extractedData = RetrieveDataFromWorkbook<fieldValueType>(workSheetName, extractionRequest);
                yield return extractedData;
            }
        }

        public IEnumerable<KeyValuePair<string, fieldValueType>> RetrieveDataCollectionFromSpecificWorksheet<fieldValueType>(int workSheetNum, IEnumerable<ExtractionRequest> extractionRequests)
        {
            foreach (var extractionRequest in extractionRequests)
            {
                var extractedData = RetrieveDataFromWorkbook<fieldValueType>(workSheetNum, extractionRequest);
                yield return extractedData;
            }
        }

        public KeyValuePair<string, fieldValueType> RetrieveDataFromWorkbook<fieldValueType>(string workSheetName, string extractedDataFieldName, int rowNum, int columnNum)
        {
            if (string.IsNullOrWhiteSpace(extractedDataFieldName)) throw new ArgumentException("An extraction request fieldname cannot be null or empty");

            bool worksheetExists = RequiredWorkbook.TryGetWorksheet(workSheetName, out IXLWorksheet requiredWorksheet);
            if (!worksheetExists) throw new NullReferenceException($"Worksheet {workSheetName} does not exist in {Path.GetFileName(requiredWorkbookPath)}");

            var worksheet = RequiredWorkbook.Worksheet(workSheetName);
            return GenerateExtractedData<fieldValueType>(worksheet, extractedDataFieldName, rowNum, columnNum);

        }

        public KeyValuePair<string, fieldValueType> RetrieveDataFromWorkbook<fieldValueType>(string workSheetName, ExtractionRequest extractionRequest)
        {
            return RetrieveDataFromWorkbook<fieldValueType>(workSheetName, extractionRequest.FieldName, extractionRequest.RowNum, extractionRequest.ColNum);
        }

        public KeyValuePair<string, fieldValueType> RetrieveDataFromWorkbook<fieldValueType>(int workSheetNum, string extractedDataFieldName, int rowNum, int columnNum)
        {
            if (string.IsNullOrWhiteSpace(extractedDataFieldName)) throw new ArgumentException("An extraction request fieldname cannot be null or empty");
            if (!Enumerable.Range(1, RequiredWorkbook.Worksheets.Count).Contains(workSheetNum)) throw new ArgumentOutOfRangeException("workSheetNum", $"Not within range of worksheets. Worksheets count: {RequiredWorkbook.Worksheets.Count}");

            var worksheet = RequiredWorkbook.Worksheet(workSheetNum);
            return GenerateExtractedData<fieldValueType>(worksheet, extractedDataFieldName, rowNum, columnNum);

        }

        public KeyValuePair<string, fieldValueType> RetrieveDataFromWorkbook<fieldValueType>(int workSheetNum, ExtractionRequest extractionRequest)
        {
            return RetrieveDataFromWorkbook<fieldValueType>(workSheetNum, extractionRequest.FieldName, extractionRequest.RowNum, extractionRequest.ColNum);
        }

        private KeyValuePair<string, fieldValueType> GenerateExtractedData<fieldValueType>(IXLWorksheet worksheet, string extractedDataFieldName, int rowNum, int columnNum)
        {
            var fieldValue = ExtractDataFromWorksheet<fieldValueType>(rowNum, columnNum, worksheet);
            return new KeyValuePair<string, fieldValueType>(extractedDataFieldName, fieldValue);
        }

        private valueType ExtractDataFromWorksheet<valueType>(int rowNum, int columnNum, IXLWorksheet worksheet)
        {
            bool valueExists = worksheet.Cell(rowNum, columnNum).TryGetValue(out valueType value);
            if (!valueExists) throw new NullReferenceException($"No value in row:{rowNum} col:{columnNum} in {worksheet.Name} of type {value.GetType().Name}");
            return value;
        }

        public Tuple<DataTable, IEnumerable<ParsingError>> SheetToDataTable(IXLWorksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException("worksheet", "cannot be null");

            DataTable dt = new DataTable();
            List<ParsingError> parsingErrors = new List<ParsingError>();

            bool firstRow = true;
            foreach (IXLRow row in worksheet.Rows())
            {
                // First row used as table column headers
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dt.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    if (!row.IsEmpty())
                    {
                        dt.Rows.Add();
                        int i = 0;

                        var firstCellColumnNum = row.FirstCellUsed();
                        var lastCellColumnNum = row.LastCellUsed();

                        if (firstCellColumnNum != null && lastCellColumnNum != null)
                        {
                            foreach (IXLCell cell in row.Cells(firstCellColumnNum.Address.ColumnNumber, lastCellColumnNum.Address.ColumnNumber))
                            {
                                string cellValue = "";
                                try
                                {
                                    if (cell.Value != null) cellValue = cell.Value.ToString();
                                }
                                catch(ExpressionParseException e)
                                {
                                    parsingErrors.Add(new ParsingError($"Invalid expression at {cell.Address.ToString()} - likely from a mistyped calculation", e));
                                }
                                catch (NullReferenceException e)
                                {
                                    parsingErrors.Add(new ParsingError($"null reference error at {cell.Address.ToString()} - likely from a mistyped calculation", e));

                                }

                                dt.Rows[dt.Rows.Count - 1][i] = cellValue;
                                i++;
                            }
                        }
                    }
                }
            }

            return new Tuple<DataTable, IEnumerable<ParsingError>>(dt, parsingErrors);

        }
    }
}
