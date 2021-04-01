using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using System.IO;
using System.Reflection;
namespace ConsoleApp1
{
    class Program
    {
        static void Main(string[] args)

        {
            System.Data.DataTable dataTable = new System.Data.DataTable();
            dataTable.Columns.Add("Page URL");
            dataTable.Columns.Add("Site Domain");
            dataTable.Columns.Add("Country");
            dataTable.Columns.Add("Product Name");
            dataTable.Columns.Add("Product Category");
            dataTable.Columns.Add("Quantity");
            dataTable.Columns.Add("Code");
            dataTable.Columns.Add("Currency");
            dataTable.Columns.Add("Promo Flag");
            dataTable.Columns.Add("Regular Price");
            dataTable.Columns.Add("Dicount Price");
            dataTable.Columns.Add("Dicount %");
            dataTable.Columns.Add("Date");
            using (WebClient webClient = new WebClient())
            {
                webClient.Headers.Add("x-api-key", "IuimuMneIKJd3tapno2Ag1c1WcAES97j");
                for (int k = 0; k <= 56; k++)
                {
                    string downloadString = webClient.DownloadString("https://apijumboweb.smdigital.cl/catalog/api/v2/products/search/vinos-cervezas-y-licores?page="+k);

                    //HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(downloadString);
                    //Console.WriteLine(downloadString);
                    HtmlAgilityPack.HtmlDocument docr = new HtmlAgilityPack.HtmlDocument();
                    docr.LoadHtml(downloadString);
                    Console.WriteLine(docr.Text);
                    string filecon = docr.Text;
                    MatchCollection blks = Regex.Matches(filecon, "productId[\\w\\W]*?\\]\\}");
                    foreach (var blk in blks)
                    {
                        Console.WriteLine(blk);
                        Match code = Regex.Match(blk.ToString(), "ref_id\\W+(\\d+)");
                        Match name = Regex.Match(blk.ToString(), "productName\\W+([\\w\\W]*?)\\\"\\s*\\,");
                        Match link = Regex.Match(blk.ToString(), "linkText\\W+([\\w\\W]*?)\\\"\\s*\\,");
                        Match o_price = Regex.Match(blk.ToString(), "Price\\W+(\\d+)");
                        Match p_price = Regex.Match(blk.ToString(), "PriceWithoutDiscount\\W+(\\d+)");
                        Match catagories = Regex.Match(blk.ToString(), "categories\\W+[\\w\\W]*?\\,\\s*\\\"\\/Vinos[\\w\\W]*?\\/([\\w\\W]*?)\\/");
                        Match Quantity = Regex.Match(name.ToString(), "(\\d+\\s*(?:ML|ml|Ml|CC|cc|KG|Kg|kg|L|l|MG|Mg|mg))");
                        DateTime dateTime = DateTime.UtcNow.Date;
                        string date = dateTime.ToString("dd/MM/yyyy");
                        Console.WriteLine(name.Groups[1].ToString());
                        Console.WriteLine(code.Groups[1].ToString());
                        Console.WriteLine(link.Groups[1].ToString());
                        Console.WriteLine(o_price.Groups[1].ToString());
                        Console.WriteLine(p_price.Groups[1].ToString());
                        Console.WriteLine(catagories.Groups[1].ToString());
                        Console.WriteLine(Quantity.Groups[1].ToString());
                        int percentage1 = (int.Parse(p_price.Groups[1].ToString()) - int.Parse(o_price.Groups[1].ToString())) ;
                        double per1 = percentage1 / int.Parse(p_price.Groups[1].ToString());
                        double percentage_vales = per1 * 100;
                        DataRow rows = dataTable.NewRow();
                        rows["Page URL"] = "https://www.jumbo.cl/" + link.Groups[1].ToString();
                        rows["Site Domain"] = "Chl";
                        rows["Country"] = "Spain";
                        rows["Product Name"] = name.Groups[1].ToString().Trim();
                        rows["Product Category"] = catagories.Groups[1].ToString();
                        rows["Quantity"] = Quantity.Groups[1].ToString();
                        rows["Code"] = code.Groups[1].ToString();
                        rows["Currency"] = "$(Dollers)";
                        rows["Promo Flag"] = "";
                        rows["Regular Price"] = o_price.Groups[1].ToString();
                        rows["Dicount Price"] = p_price.Groups[1].ToString();
                        rows["Dicount %"] = percentage_vales + "%";
                        rows["Date"] = date;
                        dataTable.Rows.Add(rows);
                    }
                }

                string folderpath = Directory.GetCurrentDirectory();
                if (!Directory.Exists(folderpath))
                {
                    Directory.CreateDirectory(folderpath);
                }
                using (XLWorkbook wb = new XLWorkbook())
                {
                    wb.Worksheets.Add(dataTable, "Datas");
                    wb.SaveAs(folderpath + "Web_Crawling_Output.xlsx");

                }
            }
        }
    }
}
