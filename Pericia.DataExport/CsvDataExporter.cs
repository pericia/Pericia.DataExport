using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace Pericia.DataExport
{
    public class CsvDataExporter : DataExporter
    {
        private StreamWriter writer;

        private List<string> currentLine;

        private const char SEPARATOR = ';';
        private const string QUOTE = "\"";
        private const string ESCAPED_QUOTE = "\"\"";
        private static readonly char[] CHARACTERS_THAT_MUST_BE_QUOTED = { SEPARATOR, '"', '\n' };


        public CsvDataExporter()
        {
            writer = new StreamWriter(stream);
            currentLine = new List<string>();
        }

        protected override void NewLine()
        {
            writer.WriteLine(String.Join(SEPARATOR.ToString(CultureInfo.InvariantCulture), currentLine));
            currentLine = new List<string>();
        }

        protected override void WriteData(object data)
        {
            currentLine.Add(Escape(data));
        }
        

        private static string Escape(object o)
        {
            if (o == null)
            {
                return "";
            }

            var s = o.ToString();

            if (s.Contains(QUOTE))
            {
                s = s.Replace(QUOTE, ESCAPED_QUOTE);
            }

            if (s.IndexOfAny(CHARACTERS_THAT_MUST_BE_QUOTED) > -1)
            {
                s = QUOTE + s + QUOTE;
            }

            return s;
        }

        bool csvStarted = false;
        protected override void NewSheet(string? name)
        {
            if (csvStarted)
            {
                throw new NotSupportedException("You can't add several sheets to a csv file");
            }
            csvStarted = true;
        }

        public override MemoryStream GetFile()
        {
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            writer.Dispose();
        }
    }
}
