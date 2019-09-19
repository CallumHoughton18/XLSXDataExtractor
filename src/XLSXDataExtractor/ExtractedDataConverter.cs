using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using XLSXDataExtractor.Common;
using XLSXDataExtractor.Models;

namespace XLSXDataExtractor
{
    public static class ExtractedDataConverter
    {
        public static IXLWorksheet ConvertToWorksheet(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData)
        {
            if (TwoDiColOfExtractedData == null) throw new ArgumentNullException("TwoDiColOfExtractedData", "Cannot be null");

            string sheetName = "newsheet";
            var genDataTable = GenerateDataTable(TwoDiColOfExtractedData);
            var xlWorkbook = new XLWorkbook();
            xlWorkbook.AddWorksheet(genDataTable, sheetName);
            return xlWorkbook.Worksheet(sheetName);
        }

        public static IXLWorksheet ConvertToWorksheet(DataTable extractedDataTable)
        {
            if (extractedDataTable == null) throw new ArgumentNullException("extractedDataTable", "Cannot be null");

            string sheetName = "newsheet";
            var xlWorkbook = new XLWorkbook();
            xlWorkbook.AddWorksheet(extractedDataTable, sheetName);
            return xlWorkbook.Worksheet(sheetName);
        }

        /// <summary>
        /// Converts a given ienumerable of ienumerable of KeyValuePair objects to a CSV string
        /// </summary>
        /// <param name="extractedDataTable"></param>
        /// <param name="escapeValues">If true, csv values are surrounded by quotation marks</param>
        /// <returns>a csv string</returns>
        public static string ConvertToCSV(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData, bool escapeValues = true)
        {
            return ConvertToDelimitedString(TwoDiColOfExtractedData,",", escapeValues);
        }

        /// <summary>
        /// Converts a given datatable to a CSV string
        /// </summary>
        /// <param name="extractedDataTable"></param>
        /// <param name="escapeValues">If true, csv values are surrounded by quotation marks</param>
        /// <returns>a csv string</returns>
        public static string ConvertToCSV(DataTable extractedDataTable, bool escapeValues = true)
        {
            return ConvertToDelimitedString(extractedDataTable, ",", escapeValues);

        }

        /// <summary>
        /// Converts a collection of collections of key value pairs to a delimited string using a given delimiter
        /// </summary>
        /// <param name="TwoDiColOfExtractedData"></param>
        /// <param name="delimiter"></param>
        /// <param name="escapeValues">If true, values are surrounded with quotation marks.</param>
        /// <returns>a delimited string</returns>
        public static string ConvertToDelimitedString(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData, string delimiter, bool escapeValues = true)
        {
            if (TwoDiColOfExtractedData == null) throw new ArgumentNullException("TwoDiColOfExtractedData", "Cannot be null");

            var genDataTable = GenerateDataTable(TwoDiColOfExtractedData);

            string csvText = genDataTable.ToDelimitedString(escapeValues, delimiter);
            return csvText;
        }

        /// <summary>
        /// Converts a datatable to a delimited string using a given delimiter
        /// </summary>
        /// <param name="TwoDiColOfExtractedData"></param>
        /// <param name="delimiter"></param>
        /// <param name="escapeValues">If true, values are surrounded with quotation marks.</param>
        /// <returns>a delimited string</returns>
        public static string ConvertToDelimitedString(DataTable extractedDataTable, string delimiter, bool escapeValues = true)
        {
            if (extractedDataTable == null) throw new ArgumentNullException("extractedDataTable", "Cannot be null");

            string csvText = extractedDataTable.ToDelimitedString(escapeValues, delimiter);
            return csvText;
        }

        public static DataTable GenerateDataTable(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData)
        {
            var dataTable = new DataTable();

            var dataTableHeaders = TwoDiColOfExtractedData.SelectMany(x => x.Select(y => y.Key)).Distinct().Select(x => new DataColumn(x));
            dataTable.Columns.AddRange(dataTableHeaders.ToArray());

            int i = 0;
            foreach (var extractedDataCollection in TwoDiColOfExtractedData)
            {
                var currentRow = dataTable.NewRow();


                int j = 0;
                foreach (var extractedData in extractedDataCollection)
                {
                    try
                    {
                        currentRow[extractedData.Key] = extractedData.Value;
                    }
                    catch (ArgumentNullException e)
                    {
                        throw new ArgumentNullException("extractedData", $"An ExtractedData object at position {i},{j} has a null field name, null field names cannot be added.");
                    }
                    j++;
                }
                i++;

                dataTable.Rows.Add(currentRow);
            }

            return dataTable;
        }
    }
}
