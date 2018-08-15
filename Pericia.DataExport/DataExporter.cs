using Pericia.DataExport.Exporters;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Pericia.DataExport
{
    public class DataExporter
    {
        private IFormatExporter _exporter;
        public DataExporter(ExportFormat format)
        {
            switch (format)
            {
                case ExportFormat.Csv:
                    _exporter = new CsvExporter();
                    break;
                default:
                    throw new NotImplementedException();
            }
        }

        public Stream Export<T>(IEnumerable<T> data)
        {
            var typeInfo = typeof(T).GetTypeInfo();

            var properties = new List<(PropertyInfo Prop, ExportColumnAttribute Attr)>();
            foreach (var prop in typeInfo.GetProperties())
            {
                var attribute = prop.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    properties.Add((prop, attribute));
                }
            }

            properties = properties.OrderBy(a => a.Attr.Order).ToList();

            // Write headers
            foreach (var prop in properties)
            {
                _exporter.WriteData(prop.Attr.Title);
            }
            _exporter.NewLine();

            foreach (var line in data)
            {
                foreach (var prop in properties)
                {
                    _exporter.WriteData(prop.Prop.GetValue(line).ToString());
                }
                _exporter.NewLine();
            }

            return _exporter.GetStream();
        }
    }
}
