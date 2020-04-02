using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;

namespace TWFututrInfo
{
    class Program
    {
        static void Main(string[] args)
        {
            #region 計算台灣期貨 今日  是否為  結算日期
            DateTime futureEndDate = GetFutureDate();
            string sMonth = string.Empty;
            //所以結算前一天 需要轉檔的是....  月份+1的交易資料
            //結算當天 會直接交易 下一個月的商品
            if (DateTime.Today.Date == futureEndDate.Date.AddDays(-1).Date)
            {
                sMonth = DateTime.Now.AddMonths(1).Month.ToString("00");
            }
            else
            {
                sMonth = DateTime.Now.Month.ToString("00");
            }
            HtmlWeb client = new HtmlWeb();
            //client.
            HtmlAgilityPack.HtmlDocument doc = client.Load(@"http://info512.taifex.com.tw/Future/FusaQuote_Norl.aspx");
            HtmlNodeCollection allFortuneContentNodes = doc.DocumentNode.SelectNodes(@"//table[@class='custDataGrid']/tr");
            DateTime dtFutureEnd = GetFutureDate();
            int indexMTX = 0;
            foreach (HtmlNode node in allFortuneContentNodes)
            {
                List<HtmlNode> temp01 = node.Elements("td").ToList();
                if (temp01.Any(i => i.InnerText.Contains("小臺指期" + sMonth)))
                {
                    string sOpenP = temp01[10].InnerText.Trim(); //開盤價               
                    string sTodayHeight = temp01[11].InnerText.Trim(); //開高價               
                    string sTodayLow = temp01[12].InnerText.Trim(); //開低價               
                    string sUporLowP = temp01[7].InnerText.Trim(); //漲跌幅               
                    string sYesterDayP = temp01[13].InnerText.Trim(); //漲跌幅               
                    //  openpriceBox.Text = sOpenPriece;
                    // UpOrLowpriceBox.Text = sUporLow;
                    break;
                }
                indexMTX++;
            }

            #endregion


            #region 在這邊找到期貨歷史資料，用FTP下載 放到D...資料夾

            //在這邊找到期貨歷史資料，用FTP下載 放到D...資料夾
            //開始做轉換寫入，這邊的作法 有用到 sql大量寫入類別
            //  http://www.coco-in.net/forum-56-1.html
            string[] fileEntries = Directory.GetFiles(@"D:\FutureData");
            foreach (var item in fileEntries)
            {
                if (item.EndsWith("rpt"))
                {
                    string[] sDateArray = item.Split('_');
                    DateTime tempDate = Convert.ToDateTime(string.Format("{0}/{1}/{2}", sDateArray[1], sDateArray[2], sDateArray[3].Replace(".rpt", "")));
                    string sMonth = tempDate.ToString("yyyyMM");
                    string sDate = tempDate.ToString("yyyy/MM/dd");
                    string[] txtArray = File.ReadAllLines(item);
                    List<string> tempArray = new List<string>();
                    tempArray = txtArray.Where(i => i.Contains("MTX") && i.Contains(sMonth)).ToList();
                    if (txtArray.Count() < 1)
                    {
                        tempDate = tempDate.AddMonths(1);
                        sMonth = tempDate.ToString("yyyyMM");
                        tempArray = txtArray.Where(i => i.Contains("MTX") && i.Contains(sMonth)).ToList();
                    }
                    var dt = new DataTable();
                    dt.Columns.Add("temp01", typeof(string));
                    dt.Columns.Add("temp02", typeof(string));
                    dt.Columns.Add("temp03", typeof(string));
                    dt.Columns.Add("temp04", typeof(string));
                    dt.Columns.Add("temp05", typeof(string));
                    dt.Columns.Add("temp06", typeof(string));
                    dt.Columns.Add("temp07", typeof(string));
                    foreach (var value in tempArray)
                    {
                        var row = dt.NewRow();
                        string[] futuredata = value.Split(',');
                        string tempMonth = futuredata[2].Trim();
                        if (tempMonth != sMonth) continue;
                        row["temp01"] = futuredata[0].Trim();
                        row["temp02"] = futuredata[1].Trim();
                        row["temp03"] = futuredata[2].Trim();
                        row["temp04"] = futuredata[3].Trim();
                        row["temp05"] = futuredata[4].Trim();
                        row["temp06"] = futuredata[5].Trim();
                        //DateTime dttemp = DateTime.ParseExact("20140804133047", "yyyyMMddHHmmss", System.Globalization.CultureInfo.CurrentCulture);
                        row["temp07"] = sDate;
                        dt.Rows.Add(row);
                    }
                    //大量寫入
                    using (TransactionScope tx = new TransactionScope())
                    {
                        using (SqlConnection cn = new SqlConnection("server=127.0.0.1;uid=sa;pwd=321456852;database=dbstock"))
                        {
                            cn.Open();
                            using (SqlBulkCopy sb = new SqlBulkCopy(cn))
                            {
                                sb.DestinationTableName = "dbo.tempFuture";
                                sb.WriteToServer(dt);
                            }
                        }
                        tx.Complete();
                    }
                }
            }
            #endregion

        }


        #region method 
        /// <summary>
        /// 取得結算日期
        /// </summary>
        /// <returns></returns>
        static DateTime GetFutureDate()
        {
            DateTime mothFirst = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            DateTime firstweek = mothFirst.AddDays(1 - (int)mothFirst.DayOfWeek);
            int temp = (int)mothFirst.DayOfWeek;


            DateTime dtFutureEnd;

            if (temp <= 3)
            {
                dtFutureEnd = firstweek.AddDays(2).AddDays(14);
            }
            else
            {
                dtFutureEnd = firstweek.AddDays(2).AddDays(21);
            }

            return dtFutureEnd;
        }

        #endregion
    }
}
