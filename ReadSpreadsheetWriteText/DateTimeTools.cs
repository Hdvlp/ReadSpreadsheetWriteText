using System;

public class DateTimeTools
{
    public static String GenerateDateTime()
    {
        DateTime dateTime = DateTime.Now;
        string dateTimeLine = dateTime.ToString("yyyyMMdd'_'HHmmss'_'zzz");
        dateTimeLine = dateTimeLine.Replace(":", "");
        dateTimeLine = dateTimeLine.Replace("-", "n");
        dateTimeLine = dateTimeLine.Replace("+", "p");
        return dateTimeLine;
    }
}