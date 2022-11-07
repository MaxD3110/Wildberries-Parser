using System;
using System.IO;
using System.Net.Http;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace WBParserAPI
{
    internal class Program
    {
        static void Main(string[] args)
        {

            try
            {
                string searchString;
                StreamReader keyFile = new StreamReader("Keys.txt");

                var executableFolderPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
                Excel.Application excelApp = new Excel.Application();

                Excel.Workbook excelWb = excelApp.Workbooks.Open(executableFolderPath + @"\example.xlsx");

                int excelPage = 1;

                while ((searchString = keyFile.ReadLine()) != null)
                {
                    Excel.Worksheet excelSheet = excelWb.Sheets[excelPage];
                    excelSheet.Cells.ClearContents();

                    var result = GetData(url: $"https://www.wildberries.by/catalog/0/search.aspx?sort=popular&search={searchString}");

                    excelSheet.Cells[1, "A"].Value = "Title";
                    excelSheet.Cells[1, "B"].Value = "Brand";
                    excelSheet.Cells[1, "C"].Value = "Id";
                    excelSheet.Cells[1, "D"].Value = "Feedbacks";
                    excelSheet.Cells[1, "E"].Value = "Price";

                    int excelRow = 2;

                    if (result != null)
                    {
                        Console.OutputEncoding = System.Text.Encoding.UTF8;

                        foreach (var row in result.data.products)
                        {
                            if(row.promoTextCat != "ДЕНЬ ШОПИНГА")
                            {
                                excelSheet.Cells[excelRow, "A"].Value = row.name;
                                excelSheet.Cells[excelRow, "B"].Value = row.brand;
                                excelSheet.Cells[excelRow, "C"].Value = row.id;
                                excelSheet.Cells[excelRow, "D"].Value = row.feedbacks;
                                excelSheet.Cells[excelRow, "E"].Value = row.priceU;

                                Console.WriteLine($"Артикул {row.id} записан!");

                                excelRow++;
                            }
                        }
                    }
                    excelPage++;
                }

                excelWb.Close(true);
                excelApp.Quit();
                keyFile.Close();

                Console.WriteLine($"Готово! Excel документ находится по адресу: {executableFolderPath}\\example.xlsx");
                Console.ReadLine();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        private static ProductCard GetData(string url)
        {
            try
            {
                using (HttpClientHandler hdl = new HttpClientHandler
                {
                    AllowAutoRedirect = false,
                    AutomaticDecompression = System.Net.DecompressionMethods.GZip | System.Net.DecompressionMethods.Deflate | System.Net.DecompressionMethods.None,
                    CookieContainer = new System.Net.CookieContainer()
                })
                {
                    using (HttpClient clnt = new HttpClient(hdl, false))
                    {
                        clnt.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 OPR/91.0.4516.106");
                        clnt.DefaultRequestHeaders.Add("Accept", "*/*");
                        clnt.DefaultRequestHeaders.Add("Accept-Language", "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7");
                        clnt.DefaultRequestHeaders.Add("Connection", "keep-alive");
                        clnt.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");

                        using (var resp = clnt.GetAsync(url).Result)
                        {
                            if (!resp.IsSuccessStatusCode)
                            {
                                return null;
                            }
                        }
                    }

                    using (HttpClient clnt = new HttpClient(hdl, false))
                    {
                        clnt.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 OPR/91.0.4516.106");
                        clnt.DefaultRequestHeaders.Add("Accept", "*/*");
                        clnt.DefaultRequestHeaders.Add("Accept-Language", "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7");
                        clnt.DefaultRequestHeaders.Add("Connection", "keep-alive");
                        clnt.DefaultRequestHeaders.Add("Referer", url);

                        using (var resp = clnt.GetAsync($"https://search.wb.ru/exactmatch/sng/common/v4/search?__tmp=by&appType=1&couponsGeo=12,7,3,21&curr=byn&dest=12358386,12358403,-70563,-8139704&emp=0&lang=ru&locale=by&pricemarginCoeff=1&query={url.Substring(url.LastIndexOf('=') + 1)}&reg=0&regions=80,83,4,33,70,82,69,68,86,30,40,48,1,22,66,31&resultset=catalog&sort=popular&spp=0&suppressSpellcheck=false").Result)
                        {
                            if (resp.IsSuccessStatusCode)
                            {
                                var json = resp.Content.ReadAsStringAsync().Result;
                                if (!string.IsNullOrEmpty(json))
                                {
                                    ProductCard result = Newtonsoft.Json.JsonConvert.DeserializeObject<ProductCard>(json);
                                    return result;
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }

            return null;
        }
    }
}
