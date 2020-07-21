using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Http;
using HtmlAgilityPack;
using System.Runtime.InteropServices;

namespace Product_Tracker
{    
    class WebScraper
    {
        static List<string> webUrls = new List<string>();
        static Dictionary<string, string> domIDs = new Dictionary<string, string>();
        static Dictionary<string, string> domTags = new Dictionary<string, string>();
        static Dictionary<string, string> results = new Dictionary<string, string>();

        public WebScraper(string path, string configSheetname)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook data_Workbook = xlApp.Workbooks.Open(path);

            Excel._Worksheet config_Worksheet = data_Workbook.Sheets[configSheetname];
            Excel.Range config_xlRange = config_Worksheet.UsedRange;

            for (int r = 2; r <= config_xlRange.Rows.Count; r++)
            {
                if (config_xlRange.Cells[r, 9] != null && config_xlRange.Cells[r, 9].Value2 != null)
                {
                    webUrls.Add(config_xlRange.Cells[r, 9].Value);
                }

                if (config_xlRange.Cells[r, 6] != null
                    && config_xlRange.Cells[r, 6].Value2 != null
                    && config_xlRange.Cells[r, 7] != null
                    && config_xlRange.Cells[r, 7].Value2 != null)
                {
                    domIDs.Add(config_xlRange.Cells[r, 6].Value, config_xlRange.Cells[r, 7].Value);
                }

                if (config_xlRange.Cells[r, 1] != null
                    && config_xlRange.Cells[r, 1].Value2 != null
                    && config_xlRange.Cells[r, 2] != null
                    && config_xlRange.Cells[r, 2].Value2 != null
                    && config_xlRange.Cells[r, 3] != null
                    && config_xlRange.Cells[r, 3].Value2 != null
                    && config_xlRange.Cells[r, 4] != null
                    && config_xlRange.Cells[r, 4].Value2 != null)
                {
                    domTags.Add(config_xlRange.Cells[r, 1].Value,
                        config_xlRange.Cells[r, 2].Value + "," +
                        config_xlRange.Cells[r, 3].Value + "," +
                        config_xlRange.Cells[r, 4].Value);
                }
            }
            data_Workbook.Close();
            Marshal.ReleaseComObject(data_Workbook);
            Marshal.ReleaseComObject(xlApp);

            data_Workbook = null;
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        public void Scrape(string path, string flipkartTrackerSheetname)
        {
            GetHtmlAsync(path, flipkartTrackerSheetname);
        }

        private static async void GetHtmlAsync(string path, string trackerSheetname)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook data_Workbook = xlApp.Workbooks.Open(path);
            Excel._Worksheet productTracker_Worksheet = data_Workbook.Sheets[trackerSheetname];
            Excel.Range productTracker_xlRange = productTracker_Worksheet.UsedRange;

            int row = Int32.Parse(productTracker_xlRange.Cells[1, 2].Value2.ToString());

            foreach (var url in webUrls)
            {
                Console.WriteLine(url);
                var httpClient = new HttpClient();
                var html = await httpClient.GetStringAsync(url);

                var htmlDocument = new HtmlDocument();
                htmlDocument.LoadHtml(html);

                foreach (var ele in domIDs)
                {
                    results.Add(ele.Key, htmlDocument.GetElementbyId(ele.Value).InnerText.Trim().Replace('\n', ' ').Replace("  ", "").ToString());
                }

                foreach (var ele in domTags)
                {
                    var product = htmlDocument.DocumentNode.Descendants(ele.Value.Split(',')[0].ToString())
                        .Where(node => node.GetAttributeValue(ele.Value.Split(',')[1].ToString(), "")
                        .Equals(ele.Value.Split(',')[2].Split(';')[0].ToString())).ToList();

                    if (product.Count == 0)
                    {
                        try
                        {
                            product = htmlDocument.DocumentNode.Descendants(ele.Value.Split(',')[0].ToString())
                            .Where(node => node.GetAttributeValue(ele.Value.Split(',')[1].ToString(), "")
                            .Equals(ele.Value.Split(',')[2].Split(';')[1].ToString())).ToList();
                        }catch { }
                    }

                    switch (product.Count)
                    {
                        case 0:
                            results.Add(ele.Key, ele.Key + " Not Found");
                            break;
                        case 1:
                            results.Add(ele.Key, product[0].InnerText.ToString().Trim().Replace('\n', ' ').Replace("  ", ""));
                            break;
                        default:
                            results.Add(ele.Key, "More than one " + ele.Key + " found");
                            break;
                    }
                }

                int col = 1;

                productTracker_xlRange.Cells[row, col].Value2 = DateTime.Now.ToString("M/d/yyyy");
                col++;

                productTracker_xlRange.Cells[row, col].Value2 = DateTime.Now.ToString("HH:mm:ss tt");
                col++;

                foreach (var ele in results)
                {
                    productTracker_xlRange.Cells[row, col].Value2 = ele.Value;
                    col++;
                }
                row++;
                results.Clear();
            }

            productTracker_xlRange.Cells[1, 2].Value2 = row;
            Console.WriteLine("Data Extracted from Website");

            xlApp.DisplayAlerts = false;

            data_Workbook.Close(SaveChanges: true, Filename: path, Excel.XlFileAccess.xlReadWrite);

            Marshal.ReleaseComObject(data_Workbook);
            Marshal.ReleaseComObject(xlApp);

            data_Workbook = null;
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
