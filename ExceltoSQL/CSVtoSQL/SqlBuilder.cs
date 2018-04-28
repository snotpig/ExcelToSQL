using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExceltoSQL
{
    class SqlBuilder
    {
        private const string TableName = "#table";
        private readonly IEnumerable<Worksheet> _worksheets;
        private IEnumerable<IEnumerable<string>> _values;
        public List<Column> Columns { get; set; }

        public SqlBuilder(IEnumerable<Worksheet> worksheets)
        {
            _worksheets = worksheets;
            OpenWorksheet(1);
        }

        public void OpenWorksheet(int index)
        {
            var worksheet = _worksheets.ElementAt(index - 1);
            _values = worksheet.Rows.Skip(1);
            Columns = worksheet.Rows.First()
                .Select((h, i) => new Column(h, _values.All(v => (v.Count() <= i) || int.TryParse(v.ElementAt(i), out var n))
                ? "int"
                : _values.All(v => (v.Count() <= i) || decimal.TryParse(v.ElementAt(i), out var d))
                ? $"decimal(12,{_values.Max(v => v.ElementAt(i).Length - Math.Abs(v.ElementAt(i).LastIndexOf('.')) - 1).ToString()})"
                : $"nvarchar({_values.Max(v => (v.Count() > i) ? v.ElementAt(i).Length : 1).ToString()})")).ToList();
        }

        public string GetSql()
        {
            var selectedColumns = Columns.Where(c => c.Include).ToList();
            var selectedValues = _values.Select(l => Columns
                .Select((c, i) => c.Type.Substring(0, 3) == "nva" ? $"'{((l.Count() > i) ? l.ElementAt(i) : "") }'" : ((l.Count() > i) ? l.ElementAt(i) : "null"))
                .Where((v, i) => Columns.ElementAt(i).Include));
            var sql = new StringBuilder($"IF OBJECT_ID('tempdb..{TableName}') IS NOT NULL DROP TABLE {TableName}\r\nGO\r\n\r\n");
            sql.Append($"CREATE TABLE {TableName} (tblId int,{selectedColumns.Aggregate(new StringBuilder(), (sb, c) => sb.Append($"[{c.Name}] {c.Type},"))}");
            sql.Length--;
            sql.Append($")\r\nINSERT INTO #table (tblId,{string.Join(", ", selectedColumns.Select(c => $"[{c.Name}]"))})\r\nVALUES\t");
            sql.Append(string.Join(",\r\n\t\t", selectedValues.Select((l, i) => $"({i+1},{l.Aggregate((t, v) => t + $@",{v}")})")));
            sql.Append($"\r\n\r\n{selectedColumns.Aggregate(new StringBuilder(), (sb, c) => sb.Append($"DECLARE @{c.Name.Replace(' ', '_').Replace('/', '_')} {c.Type};\r\n"))}");
            sql.Append($"DECLARE @MaxId int = {_values.Count()};\r\nDECLARE @CurrentId int = 1;\r\n\r\nWHILE @CurrentId <= @MaxId\r\nBEGIN\r\n\tSELECT TOP 1");
            sql.Append($"{selectedColumns.Aggregate(new StringBuilder(), (sb, c) => sb.Append($"\r\n\t@{c.Name.Replace(' ', '_').Replace('/', '_')} = [{c.Name}],"))}");
            sql.Length--;
            sql.Append($"\tFROM #table\r\n\tWHERE TblId = @CurrentId\r\n\r\n\t--insert logic here...\r\n\r\n\tSELECT @CurrentId = @CurrentId + 1\r\nEND");
            return sql.ToString();
        }
    }
}
