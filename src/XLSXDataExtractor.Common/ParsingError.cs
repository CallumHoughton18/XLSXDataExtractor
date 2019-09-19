using System;
using System.Collections.Generic;
using System.Text;

namespace XLSXDataExtractor.Common
{
    public class ParsingError
    {
        public string ParsingErrorMsg { get; set; }
        public Exception ParsingErrorException { get; set; }

        public ParsingError(string parsingErrorMsg, Exception parsingErrorException)
        {
            ParsingErrorMsg = parsingErrorMsg;
            ParsingErrorException = parsingErrorException;
        }
    }
}
