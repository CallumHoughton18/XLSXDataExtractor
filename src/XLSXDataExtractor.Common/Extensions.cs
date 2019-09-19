using System;
using System.Data;
using System.Text;

namespace XLSXDataExtractor.Common
{
    public static class Extensions
    {
        public static string ToCSVString(this DataTable dataTable, bool escapeValues = true)
        {
            return ToDelimitedString(dataTable, escapeValues, ",");
        }

        public static string ToDelimitedString(this DataTable dataTable, bool escapeValues = true, string delimiter = ",")
        {
            StringBuilder stringBuilder = new StringBuilder();

            if (dataTable.Columns.Count == 0)
                return null;

            foreach (var column in dataTable.Columns)
            {
                if (column == null) stringBuilder.Append(delimiter);
                else if (escapeValues) stringBuilder.Append($"\"{column.ToString().Replace("\"", "\"\"")}\"{delimiter}");
                else stringBuilder.Append($"{column.ToString()}{delimiter}");
            }

            stringBuilder.Replace(delimiter, Environment.NewLine, stringBuilder.Length - 1, 1);

            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    if (column == null) stringBuilder.Append(delimiter);
                    else if (escapeValues) stringBuilder.Append($"\"{column.ToString().Replace("\"", "\"\"")}\"{delimiter}");
                    else stringBuilder.Append($"{column.ToString()}{delimiter}");
                }
                stringBuilder.Replace(delimiter, Environment.NewLine, stringBuilder.Length - 1, 1);
            }

            return stringBuilder.ToString().TrimEnd(Environment.NewLine.ToCharArray());
        }
    }
}
