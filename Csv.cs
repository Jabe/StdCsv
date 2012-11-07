using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace StdCsv
{
    public class Csv
    {
        public Csv()
        {
            // defaults picked to open in excel 2013 right away

            FieldDelimiter = ";";
            FieldQuote = "\"";
            LineEnd = "\r\n";

            Culture = CultureInfo.InvariantCulture;
        }

        public string FieldDelimiter { get; set; }
        public string FieldQuote { get; set; }
        public string LineEnd { get; set; }
        public string NullSubstitution { get; set; }
        public string NewLineSubstitution { get; set; }
        public bool QuoteAllFields { get; set; }
        public bool SortColumns { get; set; }
        public CultureInfo Culture { get; set; }

        public async Task WriteDictionary(IEnumerable<IDictionary<string, object>> rows, TextWriter writer,
                                          bool includeHeader = true)
        {
            Tuple<string, Func<object, object>>[] schema = null;

            foreach (var row in rows)
            {
                if (schema == null)
                {
                    schema = GetSchema(row).ToArray();

                    if (includeHeader)
                    {
                        await WriteHeader(schema, writer);
                    }
                }

                await WriteRow(row, schema, writer);
            }

            await writer.FlushAsync();
        }

        public async Task WriteObjects<T>(IEnumerable<T> rows, TextWriter writer, bool includeHeader = true)
        {
            Tuple<string, Func<object, object>>[] schema = GetSchema(typeof (T)).ToArray();

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

        private IEnumerable<Tuple<string, Func<object, object>>> GetSchema(IDictionary<string, object> row)
        {
            foreach (var kv in row)
            {
                string key = kv.Key;
                Func<object, object> accessor = r => ((IDictionary<string, object>) r)[key];
                yield return Tuple.Create(key, accessor);
            }
        }

        private IEnumerable<Tuple<string, Func<object, object>>> GetSchema(Type type)
        {
            IEnumerable<MemberInfo> info = type.GetMembers()
                .Where(x => x.MemberType == MemberTypes.Field || x.MemberType == MemberTypes.Property);

            if (SortColumns)
                info = info.OrderBy(x => x.Name);

            foreach (MemberInfo member in info)
            {
                var prop = member as PropertyInfo;
                var field = member as FieldInfo;

                Func<object, object> accessor = row => null;

                if (prop != null)
                    accessor = row => prop.GetValue(row, null);
                else if (field != null)
                    accessor = field.GetValue;

                yield return Tuple.Create(member.Name, accessor);
            }
        }

        private async Task WriteHeader(ICollection<Tuple<string, Func<object, object>>> schema, TextWriter writer)
        {
            foreach (var member in schema)
            {
                string str = MakeField(member.Item1);

                if (!ReferenceEquals(member, schema.Last()))
                    str += FieldDelimiter;

                await writer.WriteAsync(str);
            }

            await writer.WriteAsync(LineEnd);
        }

        private async Task WriteRow<T>(T row, ICollection<Tuple<string, Func<object, object>>> schema, TextWriter writer)
        {
            foreach (var member in schema)
            {
                object value = member.Item2(row);

                string str = MakeField(ConvertValue(value));

                if (!ReferenceEquals(member, schema.Last()))
                    str += FieldDelimiter;

                await writer.WriteAsync(str);
            }

            await writer.WriteAsync(LineEnd);
        }

        private string MakeField(string value)
        {
            if (value == null)
            {
                value = NullSubstitution;
            }

            if (value == null)
            {
                return QuoteAllFields ? FieldQuote + FieldQuote : value;
            }

            if (NewLineSubstitution != null)
            {
                value = Regex.Replace(value, Regex.Escape(LineEnd) + @"|\r\n|\n", NewLineSubstitution);
            }

            bool containsQuote = value.Contains(FieldQuote);

            bool needsQuote = containsQuote ||
                              QuoteAllFields ||
                              value.Contains(FieldDelimiter) ||
                              value.Contains('\n') ||
                              value.Contains("\r\n") ||
                              value.Contains(LineEnd);

            if (needsQuote)
            {
                if (containsQuote)
                {
                    value = value.Replace(FieldQuote, FieldQuote + FieldQuote);
                }

                value = FieldQuote + value + FieldQuote;
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