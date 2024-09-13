using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using System.IO;
using System.IO.Pipes;

namespace ReadSpreadsheetWriteText
{
    class ReadSpreadsheetTool
    {
        private string pathToSpreadsheetXlsx {  get; set; }
        public ReadSpreadsheetTool(string spreadsheetXlsx)
        {
            this.pathToSpreadsheetXlsx = spreadsheetXlsx;
        }

        public void zipExtract(string destinationPathDirectory)
        {
            ZipFile.ExtractToDirectory(pathToSpreadsheetXlsx, destinationPathDirectory, Encoding.UTF8, true);
        }


    }
}
