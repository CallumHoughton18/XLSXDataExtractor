using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXDataExtractor.Models
{
    public class ExtractionRequest
    {
        public string FieldName { get; private set; }
        public int RowNum { get; private set; }
        public int ColNum { get; set; }

        public ExtractionRequest(string fieldName, int rowNum, int colNum)
        {
            FieldName = fieldName;
            RowNum = rowNum;
            ColNum = colNum;
        }
    }
}
