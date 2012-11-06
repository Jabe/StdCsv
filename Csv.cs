using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace StdCsv
{
    public class Csv
    {
        public Csv()
        {
            // defaults picked to open in excel 2013 right away

            Delimiter = ";";
            NewLine = "\r\n";
            Quote = "\"";

            Culture = CultureInfo.InvariantCulture;
        }

        public string Delimiter { get; set; }
        public string NewLine { get; set; }
        public string Quote { get; set; }
        public string NullSubstitution { get; set; }
        public string NewLineSubstitution { get; set; }
        public bool QuoteAllFields { get; set; }
        public bool SortColumns { get; set; }
        public CultureInfo Culture { get; set; }

        public async Task Write<T>(IEnumerable<T> rows, TextWriter writer, bool includeHeader = true)
        {
            IList<MemberInfo> schema = GetSchema(typeof (T));

            if (includeHeader)
            {
                await WriteHeader(schema, writer);
            }

            foreach (T row in rows)
            {
                await WriteRow(row, schema, writer);
            }

            await writer.FlushAsync();
        }

        private IList<MemberInfo> GetSchema(Type type)
        {
            IEnumerable<MemberInfo> info = type.GetMembers()
                .Where(x => x.MemberType == MemberTypes.Field || x.MemberType == MemberTypes.Property);

            if (SortColumns)
                info = info.OrderBy(x => x.Name);

            return info.ToList();
        }

        private async Task WriteHeader(IList<MemberInfo> schema, TextWriter writer)
        {
            foreach (MemberInfo member in schema)
            {
                string str = MakeField(member.Name);

                if (member != schema.Last())
                    str += Delimiter;

                await writer.WriteAsync(str);
            }

            await writer.WriteAsync(NewLine);
        }

        private async Task WriteRow<T>(T row, IList<MemberInfo> schema, TextWriter writer)
        {
            foreach (MemberInfo member in schema)
            {
                var prop = member as PropertyInfo;
                var field = member as FieldInfo;

                object value = null;

                if (prop != null)
                    value = prop.GetValue(row, null);
                else if (field != null)
                    value = field.GetValue(row);

                string str = MakeField(ConvertValue(value));

                if (member != schema.Last())
                    str += Delimiter;

                await writer.WriteAsync(str);
            }

            await writer.WriteAsync(NewLine);
        }

        private string MakeField(string value)
        {
            if (value == null)
            {
                value = NullSubstitution;
            }

            if (value == null)
            {
                return QuoteAllFields ? Quote + Quote : value;
            }

            if (NewLineSubstitution != null)
            {
                value = value.Replace(NewLine, NewLineSubstitution);
            }

            bool containsQuote = value.Contains(Quote);

            bool needsQuote = containsQuote ||
                              QuoteAllFields ||
                              value.Contains(Delimiter) ||
                              value.Contains(NewLine);

            if (needsQuote)
            {
                if (containsQuote)
                {
                    value = value.Replace(Quote, Quote + Quote);
                }

                value = Quote + value + Quote;
            }

            return value;
        }

        private string ConvertValue(object value)
        {
            if (value is DateTime)
            {
                return ((DateTime) value).ToString("o", Culture);
            }

            if (value is DateTimeOffset)
            {
                return ((DateTimeOffset) value).ToString("o", Culture);
            }

            return Convert.ToString(value, Culture);
        }
    }
}