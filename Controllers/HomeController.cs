using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Logging;
using worddemo.Models;

namespace worddemo.Controllers
{
    public class HomeController : Controller
    {

        private readonly ILogger<HomeController> _logger;

        private string connString;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public HomeController(IWebHostEnvironment webHostEnvironment, ILogger<HomeController> logger)
        {
            _logger = logger;
            _webHostEnvironment = webHostEnvironment;
            string dataPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            dataPath = dataPath.Substring(0, dataPath.Length - 7) + "appData\\" + "Worddemo.db";
            connString = "Data Source=" + dataPath;
        }



        public IActionResult Index()
        {

            /*string sql = "select * from word order by id desc";
            SqliteConnection conn = new SqliteConnection(connString);

            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            StringBuilder strHtml = new StringBuilder();
            //流转跳转到本页
            bool flg = false;
            string DocID = "";

            string requestID = Request.Query["ID"];
            string requestflag = Request.Query["flag"];

            if (requestID != null && requestID.Trim().Length > 0)
            {
                DocID = Request.Query["ID"];
                if (requestflag != null && requestflag.Length > 0)
                {
                    flg = true;
                }
            }

            while (dr.Read())
            {
                //流转，高亮显示流转操作的文档记录
                if (dr["ID"].ToString().Equals(DocID) && flg)
                {
                    strHtml.Append("<tr style=' background-color:#D7FFEE' onmouseover='onColor(this)' onmouseout='offColor(this)'>\n");
                    strHtml.Append("<td ><img src='images/office-1.jpg' /></td>\n");
                    strHtml.Append("<td >" + dr["Subject"] + "</td>\n");

                }
                //非流转
                else
                {
                    strHtml.Append("<tr onmouseover='onColor(this)' onmouseout='offColor(this)'>\n");
                    strHtml.Append("<td><img src='images/office-1.jpg' /></td>\n");
                    strHtml.Append("<td>" + dr["Subject"] + "</td>\n");

                }

                strHtml.Append("<td>" + DateTime.Parse(dr["SubmitTime"].ToString()).ToString("yyyy/MM/dd") + "</td>\n");

                switch (dr["Status"].ToString())
                {
                    case "在线编辑":
                        strHtml.Append(" <td colspan=4><a href = \"javascript:POBrowser.openWindow('Edit/word2?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" ><span style=' color:Blue;'>在线编辑</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=张三" + "', 'width=1200px;height=800px;');\">张三批阅</a> " +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=李四" + "', 'width=1200px;height=800px;');\" >李四批阅</a> " +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word1?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">文员清稿</a> " +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word3?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">正式发文</a></td>\n");
                        break;
                    case "张三批阅":
                        strHtml.Append(" <td colspan=4><a href = \"javascript:POBrowser.openWindow('Edit/word2?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" ><span style=' color:Green;'>在线编辑</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindowModeless('Edit/word?ID=" + dr["ID"] + "&user=zhangsan'" + ", 'width=1200px;height=800px;','');\"><span style=' color:Blue;'>张三批阅</span></a>" +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=李四'" + ", 'width=1200px;height=800px;');\" >李四批阅</a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word1?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">文员清稿</a>" +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word3?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">正式发文</a></td>\n");                                    
                        break;
                    case "李四批阅":
                        strHtml.Append(" <td colspan=4><a href = \"javascript:POBrowser.openWindow('Edit/word2?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" ><span style=' color:Green;'>在线编辑</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=张三'" + ", 'width=1200px;height=800px;');\"><span style=' color:Green;'>张三批阅</span></a>" +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=李四'" + ", 'width=1200px;height=800px;');\" ><span style=' color:Blue;'>李四批阅</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word1?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">文员清稿</a>" +
               " →<a href = \"javascript:POBrowser.openWindow('Edit/word3?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">正式发文</a></td>\n");
                        break;
                    case "文员清稿":
                        strHtml.Append(" <td colspan=4><a href = \"javascript:POBrowser.openWindow('Edit/word2?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" ><span style=' color:Green;'>在线编辑</span></a>" +
               " →<a href =  \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=张三'" + ", 'width=1200px;height=800px;');\"><span style=' color:Green;'>张三批阅</span></a>" +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=李四'" + ", 'width=1200px;height=800px;');\" ><span style=' color:Green;'>李四批阅</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word1?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\"><span style=' color:Blue;'>文员清稿</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word3?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\">正式发文</a></td>\n");
                        break;
                    case "正式发文":
                        strHtml.Append(" <td colspan=4><a href = \"javascript:POBrowser.openWindow('Edit/word2?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\" ><span style=' color:Green;'>在线编辑</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=张三'" + ", 'width=1200px;height=800px;');\"><span style=' color:Green;'>张三批阅</span></a>" +
               " →<a href = \"javascript:POBrowser.openWindow('Edit/word?ID=" + dr["ID"] + "&user=李四'" + ", 'width=1200px;height=800px;');\" ><span style=' color:Green;'>李四批阅</span></a>" +
               " → <a href =  \"javascript:POBrowser.openWindow('Edit/word1?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\"><span style=' color:Green;'>文员清稿</span></a>" +
               " → <a href = \"javascript:POBrowser.openWindow('Edit/word3?ID=" + dr["ID"] + "', 'width=1200px;height=800px;');\"><span style=' color:Blue;'>正式发文</a></span></td>\n");
                        break;
                }

                if (dr["HtmlFile"] != null && dr["HtmlFile"].ToString() != "")
                    strHtml.Append(" <td><a href='doc/" + dr["HtmlFile"].ToString() + "'><span style=' color:Green;'>Html</span></a></td>\n");
                else
                    strHtml.Append(" <td>Html</td>\n");
                strHtml.Append(" </tr>\n");
            }


            dr.Close();
            conn.Close();

            ViewBag.strHtml = strHtml;*/
            return View();
        }



        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }

}



