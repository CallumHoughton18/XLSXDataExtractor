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

        public static string ConvertToCSV(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData)
        {
            if (TwoDiColOfExtractedData == null) throw new ArgumentNullException("TwoDiColOfExtractedData", "Cannot be null");

            var genDataTable = GenerateDataTable(TwoDiColOfExtractedData);

            string csvText = genDataTable.ToCSVString();
            return csvText;
        }

        private static DataTable GenerateDataTable(IEnumerable<IEnumerable<KeyValuePair<string, object>>> TwoDiColOfExtractedData)
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
