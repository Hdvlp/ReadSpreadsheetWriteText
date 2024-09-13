
using System;
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
            string dateTimeLine = DateTimeTools.GenerateDateTime();
            StringTools stringTools = new StringTools();
            string randDirName = stringTools.GenStringWith(dateTimeLine, "_");

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

            PathTools pathTools = new PathTools();

            ReadSpreadsheetTool readSpreadsheetTool = new(pathToXlsx);

            string folderPath = pathTools.PreparePathToDirectory(pathToXlsx, randDirName, Path.DirectorySeparatorChar);
            

            int localCount = 0;
            const int maxGetUniqueName = 25;
            for (int i = 0; i < maxGetUniqueName; i++)
            {
                localCount++;

                if (localCount >= maxGetUniqueName)
                {
                    throw new InvalidOperationException("Too many directory names are the same.");
                }
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                    break;
                }
                else
                {
                    randDirName = stringTools.GenStringWith(dateTimeLine, "_");
                    folderPath = pathTools.PreparePathToDirectory(pathToXlsx, randDirName, Path.DirectorySeparatorChar);
                }
            
            }
            
            readSpreadsheetTool.zipExtract(folderPath);

            string intermediateFolders = $@"localTempCache{Path.DirectorySeparatorChar}{randDirName}{Path.DirectorySeparatorChar}xl";
            string fileName = @"sharedStrings.xml";
            string pathToXml = pathTools.PreparePathToFile(pathToXlsx, intermediateFolders, fileName);

            ReadXmlOfSharedStrings readXmlOfSharedStrings = new(pathToXml);
            readXmlOfSharedStrings.ReadXmlAsync();

            intermediateFolders = $@"localTempCache{Path.DirectorySeparatorChar}{randDirName}{Path.DirectorySeparatorChar}xl{Path.DirectorySeparatorChar}worksheets";
            fileName = @"sheet1.xml";
            pathToXml = pathTools.PreparePathToFile(pathToXlsx, intermediateFolders, fileName);

            ReadXmlOfWorksheet readXmlOfWorksheet = new(pathToXml, pathToDirectory, fileNameOfXlsx);

            readXmlOfWorksheet.setPoolOfStringKVPair(readXmlOfSharedStrings.getPoolOfStringKVPairs());

            readXmlOfWorksheet.ReadXmlAsync();

        }

    }
}
