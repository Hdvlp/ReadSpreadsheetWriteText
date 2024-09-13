using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSpreadsheetWriteText
{
    class StringTools
    {
        public string GenStringWith(string dateTimeLine, string symbol)
        {
            StringBuilder sb = new();
            sb.Append($"{dateTimeLine}{symbol}{this.GenString(3)}");
            sb.Append($"{symbol}{this.GenString(3)}");
            sb.Append($"{symbol}{this.GenString(3)}");
            sb.Append($"{symbol}{this.GenString(3)}");
            sb.Append($"{symbol}{this.GenString(3)}");
            sb.Append($"{symbol}{this.GenString(3)}");
            return sb.ToString();
        }
            public string GenString(int length)
        {
            StringBuilder sb = new();
            for (int i = 0; i < length; i++)
            {
                sb.Append(this.GenString());
            }
            return sb.ToString();
        }
        public string GenString()
        {
            StringBuilder sb = new();
            string longString = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            sb.Append(longString);

            Thread.Sleep(1);
            Random random = new Random();
            Thread.Sleep(random.Next(0, 5));
            random = new Random();
            Thread.Sleep(random.Next(0, 5));
            random = new Random();
            Thread.Sleep(random.Next(0, 5));
            random = new Random();
            Thread.Sleep(random.Next(0, 5));
            random = new Random();
            Thread.Sleep(random.Next(0, 5));
            random = new Random();
            return sb.ToString(random.Next(0, longString.Length), 1);
        }
    }
}
