using CsvHelper;
using HtmlAgilityPack;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    class Program
    {
        //site URL
        static string URI = "https://srh.bankofchina.com/search/whpj/searchen.jsp";

        //Setting start date to be the current date -2 days
        static string date = DateTime.UtcNow.ToString("yyyy-MM-dd");

        //Setting end date to be the current date in specific format 
        static string expiryDate = DateTime.Now.AddDays(-2).ToString("yyyy-MM-dd");

        //List for adding currencies
        static List<string> currencies = new List<string>();
        //List for adding data
        static List<string> elements = new List<string>();
        //List for adding headers
        static List<string> headers = new List<string>();

        static string csvPath;

        //I couldn't catch from script var with name m_nRecordCount. I wanted to get that value and divide with 20 records per page to get total number of pages
        static int pagenumber = 25; 
       
        //StringBuilder for writting data
        static StringBuilder sb = new StringBuilder();

        //method for getting all currencies
        static void GetAllCurrencies()
        {
            WebClient wc = new WebClient();
            HtmlDocument doc = new HtmlDocument();
            HtmlNode.ElementsFlags.Remove("option");
            string myParametersDefault = "erectDate=" + date + "&nothing=" + expiryDate + "&pjname=GBP";
            string HtmlResultDefault = wc.UploadString(URI, myParametersDefault);
            doc.LoadHtml(HtmlResultDefault);
            
            //Finding elements by id property
            foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//select[@id='pjname']//option"))
            {
                //Console.WriteLine(node.InnerText);
                currencies.Add(node.InnerText);

            }
            //removing first one that represent display text
            currencies.RemoveAt(0);
        }

        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("============================ WEB SCRAPE ====================================");
                Console.WriteLine("Start date : {0}", expiryDate.ToString());
                Console.WriteLine("End date : {0}", date.ToString());
                Console.WriteLine("Website : {0}", URI.ToString());
                Console.WriteLine("============================== START ====================================");

                //Adding headers in list
                headers.Add("Currency Name");
                headers.Add("Buying Rate");
                headers.Add("Cash Buying Rate");
                headers.Add("Selling Rate");
                headers.Add("Cash Selling Rate");
                headers.Add("Middle Rate");
                headers.Add("Pub Time");


                foreach (var element in headers)
                {
                    sb.Append(element.ToString());
                    sb.Append(",");
                }
                //Method for getting all available currencies
                GetAllCurrencies();
                foreach (var currency in currencies)
                {

                    for (int i = 1; i <= pagenumber; i++)
                    {
                        string myParameters = "erectDate=" + date + "&nothing=" + expiryDate + "&pjname=" + currency + "&page=" + i.ToString();

                        using (WebClient wc = new WebClient())
                        {
                            wc.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                            string HtmlResult = wc.UploadString(URI, myParameters);
                            //Console.WriteLine(HtmlResult);

                            HtmlDocument htmlDoc = new HtmlDocument();
                            htmlDoc.LoadHtml(HtmlResult);
                            IEnumerable<HtmlNode> nodes =
                              htmlDoc.DocumentNode.Descendants(0)
                             .Where(n => n.HasClass("hui12_20"));

                            foreach (var element in nodes)
                            {
                                elements.Add(element.InnerText);

                                //If WebCliennt waited too long for response display accurate message
                                if (element.InnerHtml == "soryy,wrong search word submit,please check your search word!")
                                {
                                    Console.WriteLine("Too long waiting for response ! No data for {0} page {1}", currency, i.ToString());
                                    
                                }
                                else
                                {
                                    if (element.InnerHtml == currency)
                                    {
                                        sb.AppendLine();
                                        sb.Append(element.InnerHtml);
                                        sb.Append(",");
                                    }
                                    else
                                    {
                                        sb.Append(element.InnerHtml);
                                        sb.Append(",");
                                    }
                                }
                            }
                        }
                        string directory = ConfigurationManager.AppSettings.Get("csvpath");
                        csvPath = directory + "\\" + currency + " StartDate " + date.ToString() + " Endate " + expiryDate.ToString() + ".csv";
                        System.IO.File.WriteAllText(csvPath, sb.ToString());
                        Console.WriteLine("Successfully for {0} page {1}", currency, i.ToString());
                        

                    }
                    Console.WriteLine("==============================");
                    Console.WriteLine("Scraping data done for {0} currency", currency);
                    Console.WriteLine();
                    Console.WriteLine("Check data  in directory -> {0}", csvPath);
                    Console.WriteLine("==============================");

                    

                    sb.Clear();
                    foreach (var element in headers)
                    {
                        sb.Append(element.ToString());
                        sb.Append(",");
                    }
                }
            }

            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {

                Console.WriteLine("============================== FINISH ====================================");
            }
        }
    }
}
