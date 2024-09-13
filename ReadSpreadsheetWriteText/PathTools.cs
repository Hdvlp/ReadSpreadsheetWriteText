using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSpreadsheetWriteText

{
    class PathTools
    {
        public string PreparePathToFile(string path, string intermediateFolders, string fileName)
        {
            string tmpPath = path ?? "";
            return Path.Combine(Path.GetDirectoryName(tmpPath), intermediateFolders, fileName);
        }
    }
}
