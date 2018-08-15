using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Pericia.DataExport.Exporters
{
    internal class CsvExporter : IFormatExporter
    {
        private MemoryStream stream;
        private StreamWriter writer;

        public CsvExporter()
        {
            stream = new MemoryStream();
            writer = new StreamWriter(stream);
        }

        public void NewLine()
        {
            writer.WriteLine();
        }

        public void WriteData(string data)
        {
            writer.Write(data + ";");
        }

        public Stream GetStream()
        {
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

    }
}
