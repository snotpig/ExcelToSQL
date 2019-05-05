using System;

namespace ExceltoSQL
{
    static class StringExtensions
    {
        public static string ToSqlDateString(this string str)
        {
            return new DateTime(int.Parse(str.Substring(6, 4)), int.Parse(str.Substring(3, 2)), int.Parse(str.Substring(0, 2))).ToString("dd MMM yyyy");
		}
    }
}
