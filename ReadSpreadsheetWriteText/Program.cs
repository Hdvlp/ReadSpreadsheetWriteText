
using System.IO;

namespace ReadSpreadsheetWriteText
{
    class Program
    {
        public static void Main(String[] args)
        {
            if (!args.Any()) return;

            string pathToXlsx = "";
            string pathToDirectory = "";
            string fileNameOfXlsx = "";

            pathToXlsx = args[0];

            if (args.Length > 1)
            {
                pathToDirectory = args[1];
            }
            if (args.Length == 1)
            {
                pathToDirectory = Path.GetDirectoryName(pathToXlsx);
            }

            if (!Directory.Exists(pathToDirectory)) { Console.WriteLine("Directory does not exist."); return; }

            fileNameOfXlsx = Path.GetFileName(pathToXlsx);

            if (!File.Exists(pathToXlsx)) { Console.WriteLine("File does not exist."); return; }
           
            ReadSpreadsheetTool readSpreadsheetTool = new(pathToXlsx);
            readSpreadsheetTool.zipExtract();

            PathTools pathTools = new PathTools();

            string intermediateFolders = @"localTempCache\xl";
            string fileName = @"sharedStrings.xml";
            string pathToXml = pathTools.PreparePathToFile(pathToXlsx, intermediateFolders, fileName);

            ReadXmlOfSharedStrings readXmlOfSharedStrings = new(pathToXml);
            readXmlOfSharedStrings.ReadXmlAsync();

            intermediateFolders = @"localTempCache\xl\worksheets";
            fileName = @"sheet1.xml";
            pathToXml = pathTools.PreparePathToFile(pathToXlsx, intermediateFolders, fileName);

            ReadXmlOfWorksheet readXmlOfWorksheet = new(pathToXml, pathToDirectory, fileNameOfXlsx);

            readXmlOfWorksheet.setPoolOfStringKVPair(readXmlOfSharedStrings.getPoolOfStringKVPairs());

            readXmlOfWorksheet.ReadXmlAsync();

        }

    }
}
