
using System;

namespace Product_Tracker
{
    class Program : Data
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Starting...");
            Console.WriteLine("=========================");
            Console.WriteLine("Choose any one/multiple from the Following (For multiple use delimiter - ',' Example '1,2,3'):");
            Console.WriteLine("1. Amazon");
            Console.WriteLine("2. Flipkart");
            Console.WriteLine("\n=========================");

            var userInput = Console.ReadLine().Split(',');
            var path = "";
            var configSheetname = "";
            var trackerSheetname = "";

            foreach (var i in userInput)
            {
                switch (Int16.Parse(i))
                {
                    case 1:
                        path = AMAZON_PATH;
                        configSheetname = AMAZON_CONFIG_SHEETNAME;
                        trackerSheetname = AMAZON_TRACKER_SHEETNAME;
                        break;
                    case 2:
                        path = FLIPKART_PATH;
                        configSheetname = FLIPKART_CONFIG_SHEETNAME;
                        trackerSheetname = FLIPKART_TRACKER_SHEETNAME;
                        break;
                    default:
                        Console.WriteLine("Invalid Number: " + i + " . Must be 1 or ");
                        break;
                }

                WebScraper webScraper = new WebScraper(path, configSheetname);
                webScraper.Scrape(path, trackerSheetname);
            }
            Console.ReadLine();

        }        
    }
}
