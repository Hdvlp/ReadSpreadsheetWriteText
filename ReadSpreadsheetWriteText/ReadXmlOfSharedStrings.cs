using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSpreadsheetWriteText
{
    class ReadXmlOfSharedStrings
    {
        private string pathToXml { get; set; }
        private IDictionary<string, string> poolOfStringKVPairs { get;  set; }
        public ReadXmlOfSharedStrings(string pathToXml) { 
            this.pathToXml = pathToXml;
            this.poolOfStringKVPairs = poolOfStringKVPairs;
        }

        public IDictionary<string, string> getPoolOfStringKVPairs()
        {
            return poolOfStringKVPairs;
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
                    List<string> values, indices;
                    poolOfStringKVPairs = new Dictionary<string, string>();

                    DataSet dataSet = new DataSet();

                    dataSet.ReadXml(sr, XmlReadMode.InferSchema);

                    foreach (DataTable table in dataSet.Tables)
                    {
                        string requiredTable = "si";
                        if (table.ToString() != requiredTable) continue;

                        int verifyHeadingCount = 0;

                        ArrayList headings = new ArrayList { "t", "si_Id", "sst_Id" };
                        const int numOfMatchedHeading = 3;
                        for (int i = 0; i < table.Columns.Count; ++i)
                        {
                            string heading = table.Columns[i].ColumnName.Substring(0, Math.Min(20, table.Columns[i].ColumnName.Length));
                            
                            if (i < numOfMatchedHeading && heading.Equals(headings[i])) verifyHeadingCount++;

                        }

                        if (verifyHeadingCount != numOfMatchedHeading){
                            Console.WriteLine("Error: columns are not in the expected order");
                            break;
                        }

                        values = new List<string>();
                        indices = new List<string>();

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
                                    indices.Add(row[i].ToString() ?? "");
                                }
                            }
                        }
                        for (int i = 0; i < values.Count; i++)
                        {
                            poolOfStringKVPairs.Add(new KeyValuePair<string, string>(indices[i], values[i]));
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
