namespace ExceltoSQL
{
    static class StringExtensions
    {
        public static string ToSqlDateString(this string str)
        {
            return $"{str.Substring(6, 4)}{str.Substring(3, 2)}{str.Substring(0, 2)}";
        }
    }
}
