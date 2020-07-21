using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Product_Tracker
{
    class Data
    {
        // Full Path or Relative path of where AmazonData.xlsx is located.
        public const string AMAZON_PATH = "Enter Path//AmazonData.xlsx";
        // SheetName of AmazonData.xlsx where Configs are writen for Amazon
        public const string AMAZON_CONFIG_SHEETNAME = "AmazonConfig";
        // SheetName of AmazonData.xlsx where Result should be stored
        public const string AMAZON_TRACKER_SHEETNAME = "AmazonProductTracker";

        // Full Path or Relative path of where FlipkartData.xlsx is located.
        public const string FLIPKART_PATH = "Enter Path//FlipkartData.xlsx";
        // SheetName of FlipkartData.xlsx where Configs are writen for Flipkart
        public const string FLIPKART_CONFIG_SHEETNAME = "FlipkartConfig";
        // SheetName of FlipkartData.xlsx where Result should be stored
        public const string FLIPKART_TRACKER_SHEETNAME = "FlipkartProductTracker";
    }
}

