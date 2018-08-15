using System;
using System.Collections.Generic;
using System.IO;

namespace Pericia.DataExport
{
    public class DataExporter
    {
        private ExportFormat _format;
        public DataExporter(ExportFormat format)
        {
            _format = format;
        }

        public Stream Export<T>(IEnumerable<T> data)
        {


            return null;
        }
    }
}
