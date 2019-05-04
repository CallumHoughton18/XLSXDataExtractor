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
    public class ExtractedDataConverter
    {
        public void ConvertToNewXLSX(IEnumerable<IEnumerable<ExtractedData<object>>> TwoDiColOfExtractedData, string pathToSaveTo)
        {
            var genDataTable = GenerateDataTable(TwoDiColOfExtractedData);

            var xlWorkbook = new XLWorkbook();
            xlWorkbook.AddWorksheet(genDataTable, (xlWorkbook.Worksheets.Count + 1).ToString());
            xlWorkbook.SaveAs(pathToSaveTo);
        }

        public void ConvertToCSV(IEnumerable<IEnumerable<ExtractedData<object>>> TwoDiColOfExtractedData, string pathToSaveTo)
        {
            var genDataTable = GenerateDataTable(TwoDiColOfExtractedData);

            string csvText = genDataTable.ToCSVString();

            File.WriteAllText(pathToSaveTo, csvText);
        }

        private DataTable GenerateDataTable(IEnumerable<IEnumerable<ExtractedData<object>>> TwoDiColOfExtractedData)
        {
            var dataTable = new DataTable();

            var dataTableHeaders = TwoDiColOfExtractedData.SelectMany(x => x.Select(y => y.FieldName)).Distinct().Select(x => new DataColumn(x));
            dataTable.Columns.AddRange(dataTableHeaders.ToArray());

            foreach (var extractedDataCollection in TwoDiColOfExtractedData)
            {
                var currentRow = dataTable.NewRow();

                foreach (var extractedData in extractedDataCollection)
                {
                    currentRow[extractedData.FieldName] = extractedData.FieldValue;
                }

                dataTable.Rows.Add(currentRow);
            }

            return dataTable;
        }
    }
}
