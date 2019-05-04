using System;
using System.Data;
using System.Text;

namespace XLSXDataExtractor.Common
{
    public static class Extensions
    {
        public static string ToCSVString(this DataTable dataTable)
        {
            StringBuilder stringBuilder = new StringBuilder();

            if (dataTable.Columns.Count == 0)
                return null;

            foreach (var column in dataTable.Columns)
            {
                if (column == null)
                    stringBuilder.Append(",");
                else
                    stringBuilder.Append("\"" + column.ToString().Replace("\"", "\"\"") + "\",");
            }

            stringBuilder.Replace(",", Environment.NewLine, stringBuilder.Length - 1, 1);

            foreach (DataRow dr in dataTable.Rows)
            {
                foreach (var column in dr.ItemArray)
                {
                    if (column == null)
                        stringBuilder.Append(",");
                    else
                        stringBuilder.Append("\"" + column.ToString().Replace("\"", "\"\"") + "\",");
                }
                stringBuilder.Replace(",", Environment.NewLine, stringBuilder.Length - 1, 1);
            }

            return stringBuilder.ToString();
        }
    }
}
