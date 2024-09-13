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

        public void zipExtract()
        {
            string directoryPath = prepareIntermediateFolders(pathToSpreadsheetXlsx);
            ZipFile.ExtractToDirectory(pathToSpreadsheetXlsx, directoryPath, Encoding.UTF8, true);
        }

        private string prepareIntermediateFolders(string pathToFile)
        {
            if (!File.Exists(pathToFile)) { 
                throw new DirectoryNotFoundException("File does not exist."); 
            }

            string intermediateFolders = @"localTempCache";
            string folder = Path.Combine(Path.GetDirectoryName(pathToFile), intermediateFolders);
            if (!Directory.Exists(folder))
            {
                Directory.CreateDirectory(folder);
            }
            return folder;
        }
    }
}
