using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Pericia.DataExport.Exporters
{
    internal class CsvExporter : IFormatExporter
    {
        private MemoryStream stream;
        private StreamWriter writer;

        private List<string> currentLine;

        private const char SEPARATOR = ';';
        private const string QUOTE = "\"";
        private const string ESCAPED_QUOTE = "\"\"";
        private static readonly char[] CHARACTERS_THAT_MUST_BE_QUOTED = { SEPARATOR, '"', '\n' };


        public CsvExporter()
        {
            stream = new MemoryStream();
            writer = new StreamWriter(stream);
            currentLine = new List<string>();
        }

        public void NewLine()
        {
            writer.WriteLine(String.Join(SEPARATOR.ToString(), currentLine));
            currentLine = new List<string>();
        }

        public void WriteData(string data)
        {
            currentLine.Add(Escape(data));
        }

        public Stream GetStream()
        {
            writer.Flush();
            stream.Position = 0;
            return stream;
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

    }
}
