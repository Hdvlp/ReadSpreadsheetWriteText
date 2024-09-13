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

        public string PreparePathToDirectory(string pathToFile, string randDirName,
            char pathSeparator)
        {
            if (!File.Exists(pathToFile))
            {
                throw new FileNotFoundException("File does not exist.");
            }


            string intermediateFolders = $@"localTempCache{pathSeparator}{randDirName}";
            string folder = Path.Combine(Path.GetDirectoryName(pathToFile), intermediateFolders);

            return folder;
        }
    }
}
