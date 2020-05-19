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
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "appData\\" + "exceldemo.db";
            connString = "Data Source=" + dataPath;
        }

        public IActionResult Index()
        {

            string sql = "select * from  excel order by  ID  DESC ";
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            SqliteDataReader dr = cmd.ExecuteReader();
            StringBuilder strGrid = new StringBuilder();
            if (!dr.HasRows)
            {
                strGrid.Append("<tr >\r\n");
                strGrid.Append("<td colspan='5' width='100%'  align='center'>对不起，暂时没有可以操作的文档。\r\n");
                strGrid.Append("</td></tr>\r\n");
            }
            else
            {
                while (dr.Read())
                {
                    strGrid.Append("<tr  onmouseover='onColor(this)' onmouseout='offColor(this)' >\r\n");
                    strGrid.Append("<td><img src='images/office-2.jpg' /></td>\r\n");
                    strGrid.Append("<td>" + dr["Subject"].ToString() + "</td>\r\n");
                    strGrid.Append("<td>" + DateTime.Parse(dr["SubmitTime"].ToString()).ToShortDateString() + "</td>\r\n");
                    strGrid.Append("<td>" + " <a href=\"javascript:POBrowser.openWindow('Edit/Excel?id=" + dr["ID"].ToString() + "', 'width=1200px;height=800px;');\" >在线编辑</a> <a href= \"javascript:POBrowser.openWindow('Edit/Excel2?id=" + dr["ID"].ToString() + "', 'width=1200px;height=800px;');\" >只读打开</a>" + "</td>\r\n");
                    if (dr["HtmlFile"] == DBNull.Value)
                        strGrid.Append("<td>Html</td>\r\n");
                    else
                        strGrid.Append("<td><a href=\"javascript:openHtml('" + dr["HtmlFile"].ToString() + "');\">Html</a></td>\r\n");
                    strGrid.Append("\r\n");
                }
            }
            dr.Close();
            conn.Close();
            ViewBag.strHtml = strGrid;
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



