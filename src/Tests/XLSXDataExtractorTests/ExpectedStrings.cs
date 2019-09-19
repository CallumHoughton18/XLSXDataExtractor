using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLSXDataExtractorTests
{
    public static class ExpectedStrings
    {
        // There is an issue reading in the expected csv from the usual file.readalltext method employed in all csv tests as the \t separator is doubly escaped like \\t 
        // So instead for \t delimited csv tests the expected value is gotten from here directly as a string.

        public static readonly string ExpectedCSV_TabDelimited_NotEscaped = "Test0\tTest1\tTest2\tTest3\tTest4\tTest5\tTest6\tTest7\tTest8\tTest9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9\r\n" +
    "0\t1\t2\t3\t4\t\t6\t7\t8\t9";
    }
}
