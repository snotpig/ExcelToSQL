using System;

namespace ExceltoSQL
{
    static class StringExtensions
    {
        public static string ToSqlDatetimeString(this string str)
        {
            return new DateTime(
                int.Parse(str.Substring(6, 4)), 
                int.Parse(str.Substring(3, 2)),
                int.Parse(str.Substring(0, 2)), 
                str.Length > 12 ? int.Parse(str.Substring(11, 2)) : 0,
                str.Length > 15 ? int.Parse(str.Substring(14, 2)) : 0
                , 0).ToString("yyyy-MM-dd HH:mm:ss");
		}
    }
}
