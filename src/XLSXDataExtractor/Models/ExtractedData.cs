using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXDataExtractor.Models
{
    public class ExtractedData<T> : ExtractedDataBase
    {
        public override Type Type
        {
            get { return typeof(T); }
        }
        public string FieldName { get; set; }
        public T FieldValue { get; set; }

        public ExtractedData(string fieldName, T fieldValue)
        {
            FieldName = fieldName;
            FieldValue = fieldValue;
        }
    }
}
