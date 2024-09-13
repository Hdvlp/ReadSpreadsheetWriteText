using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Data;
using System.ComponentModel;
using System.Collections;

namespace ReadSpreadsheetWriteText
{
    class ReadXmlOfWorksheet
    {
        private string pathToXml { get; set; }
        private string pathToDirectory { get; set; }
        private string fileNameOfXlsx { get; set; }
        private IDictionary<string, string> poolOfStringKVPair { get; set; }

        public ReadXmlOfWorksheet(string pathToXml, string pathToDirectory, string fileNameOfXlsx)
        {
            this.pathToXml = pathToXml;
            this.pathToDirectory = pathToDirectory;
            this.fileNameOfXlsx = fileNameOfXlsx;
        }

        public void setPoolOfStringKVPair(IDictionary<string, string> poolOfStringKVPair)
        {
            this.poolOfStringKVPair = poolOfStringKVPair;
        }

        public async void ReadXmlAsync()
        {
            await RunReadXmlAsync(pathToXml);
        }

        private async Task RunReadXmlAsync(string pathToXml)
        {
            if (!File.Exists(pathToXml))
            {
                Console.WriteLine("File does not exist.");
                return;
            }

            try
            {
                using (StreamReader sr = new StreamReader(pathToXml))
                {
                    List<string> values, cells;
                    List<bool> isString;
                    DataSet dataSet = new DataSet();

                    dataSet.ReadXml(sr, XmlReadMode.InferSchema);

                    foreach (DataTable table in dataSet.Tables)
                    {
                        string requiredTable = "c";
                        if (table.ToString() != requiredTable) continue;

                        int verifyHeadingCount = 0;

                        ArrayList headings = new ArrayList { "v", "r", "s", "t", "row_Id" };
                        const int numOfMatchedHeading = 5;
                        for (int i = 0; i < table.Columns.Count; ++i)
                        {
                            string heading = table.Columns[i].ColumnName.Substring(0, Math.Min(20, table.Columns[i].ColumnName.Length));

                            if (i < numOfMatchedHeading && heading.Equals(headings[i])) verifyHeadingCount++;

                        }

                        if (verifyHeadingCount != numOfMatchedHeading)
                        {
                            Console.WriteLine("Error: columns are not in the expected order");
                            break;
                        }

                        values = new List<string>();
                        cells = new List<string>();
                        isString = new List<bool>();

                        foreach (var row in table.AsEnumerable())
                        {
                            for (int i = 0; i < table.Columns.Count; ++i)
                            {
                                if (i == 0)
                                {
                                    values.Add(row[i].ToString() ?? "");
                                }
                                else if (i == 1)
                                {
                                    cells.Add(row[i].ToString() ?? "");
                                }
                                else if (i == 3)
                                {
                                    isString.Add((row[i].ToString() ?? "") == "s");
                                }
                            }
                        }
                        StringBuilder stringBuilder = new StringBuilder();
                        for (int i = 0; i < values.Count; i++)
                        {
                            string outputVal = isString[i] ? poolOfStringKVPair[values[i]] : values[i];

                            stringBuilder.Append(cells[i].ToString() + "|");
                            stringBuilder.Append(outputVal + "|");
                            stringBuilder.AppendLine();
                        }
                        string contentsToFile = (stringBuilder.ToString());
                        var fileName = fileNameOfXlsx + ".txt";
                        var targetTextPath = Path.Combine(pathToDirectory, fileName);
                        if (!File.Exists(targetTextPath))
                        {
                            File.WriteAllText(targetTextPath, contentsToFile, Encoding.UTF8);
                        }
                    }

                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Error? Try? Check if directory or file exists.");
            }

        }


    }
}
