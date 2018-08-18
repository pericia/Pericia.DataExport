using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Pericia.DataExport
{
    public abstract class DataExporter
    {
        protected MemoryStream stream = new MemoryStream();       

        public void AddSheet<T>(IEnumerable<T> data)
        {
            var typeInfo = typeof(T).GetTypeInfo();

            var properties = new List<ColumnInfo>();
            foreach (var prop in typeInfo.DeclaredProperties)
            {
                var attribute = prop.GetCustomAttribute<ExportColumnAttribute>();
                if (attribute != null)
                {
                    properties.Add(new ColumnInfo() { Prop = prop, Attr = attribute });
                }
            }

            NewSheet();

            properties = properties.OrderBy(a => a.Attr.Order).ToList();
            // Write headers
            foreach (var prop in properties)
            {
                WriteData(prop.Attr.Title);
            }
            NewLine();

            foreach (var line in data)
            {
                foreach (var prop in properties)
                {
                    WriteData(prop.Prop.GetValue(line).ToString());
                }
                NewLine();
            }
        }


        public MemoryStream Export<T>(IEnumerable<T> data)
        {
            AddSheet(data);

            return GetFile();
        }

        public abstract MemoryStream GetFile();



        protected abstract void NewSheet();
        protected abstract void NewLine();
        protected abstract void WriteData(string data);

        private class ColumnInfo
        {
            internal PropertyInfo Prop { get; set; }
            internal ExportColumnAttribute Attr { get; set; }
        }
    }
}
