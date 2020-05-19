using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.Sqlite;

namespace exceldemo.Controllers.Edit
{
    public class EditController : Controller
    {
        private string connString;

        private readonly IWebHostEnvironment _webHostEnvironment;

        public EditController(IWebHostEnvironment webHostEnvironment)
        {
            _webHostEnvironment = webHostEnvironment;
            string rootPath = _webHostEnvironment.WebRootPath.Replace("/", "\\");
            string dataPath = rootPath.Substring(0, rootPath.Length - 7) + "appData\\" + "exceldemo.db";
            connString = "Data Source=" + dataPath;
        }


        public IActionResult excel()
        {
            string DocID = Request.Query["ID"];
            string sql = "select * from excel where ID = " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string docFile = "";
            while (dr.Read())
            {
                docFile = dr["FileName"].ToString();
            }
            dr.Close();
            conn.Close();
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";


            pageofficeCtrl.CustomMenuCaption = "自定义菜单(&N)";
            pageofficeCtrl.AddCustomMenuItem("显示标题(&T)", "CustomMenuItem1_Click()", true);
            pageofficeCtrl.AddCustomMenuItem("-", "", false);
            pageofficeCtrl.AddCustomMenuItem("领导圈阅(&D)", "CustomMenuItem2_Click()", true);

            pageofficeCtrl.AddCustomToolButton("保存", "CustomToolBar_Save()", 1);
            pageofficeCtrl.AddCustomToolButton("另存为...", "CustomToolBar_SaveAs()", 1);
            pageofficeCtrl.AddCustomToolButton("另存为Html", "CustomToolBar_SaveAsHtml()", 1);
            pageofficeCtrl.AddCustomToolButton("插入印章", "CustomToolBar_InsertSeal()", 2);
            pageofficeCtrl.AddCustomToolButton("领导圈阅", "CustomToolBar_HandDraw()", 3);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "CustomToolBar_FullScreen()", 4);
            pageofficeCtrl.BorderStyle = PageOfficeNetCore.BorderStyleType.BorderThin;

            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + docFile, PageOfficeNetCore.OpenModeType.xlsNormalEdit, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.time = DateTime.Now.ToString("yyyy年MM月dd日 dddd");
            return View();

        }

        public IActionResult excel2()
        {
            string DocID = Request.Query["ID"];
            string sql = "select * from excel where ID = " + DocID + ";";
            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;

            SqliteDataReader dr = cmd.ExecuteReader();
            PageOfficeNetCore.PageOfficeCtrl pageofficeCtrl = new PageOfficeNetCore.PageOfficeCtrl(Request);
            pageofficeCtrl.ServerPage = "../PageOffice/POServer";
            string docFile = "";
            string docSubject = "";
            while (dr.Read())
            {
                docFile = dr["FileName"].ToString();
                docSubject = dr["Subject"].ToString();
            }
            dr.Close();
            conn.Close();
            //设置保存页面
            pageofficeCtrl.SaveFilePage = "/Edit/SaveDoc";
            PageOfficeNetCore.ExcelWriter.Workbook wb = new PageOfficeNetCore.ExcelWriter.Workbook();
            wb.DisableSheetDoubleClick = true;
            wb.DisableSheetRightClick = true;
            pageofficeCtrl.SetWriter(wb);
            pageofficeCtrl.AddCustomToolButton("另存为...", "saveAs", 1);
            pageofficeCtrl.AddCustomToolButton("打印设置", "setPrint", 0);
            pageofficeCtrl.AddCustomToolButton("全屏/还原", "setFullScreen", 4);
            pageofficeCtrl.Caption = "Excel文件:" + docSubject;

            //打开Word文档
            pageofficeCtrl.WebOpen("/doc/" + docFile, PageOfficeNetCore.OpenModeType.xlsReadOnly, "tom");
            ViewBag.POCtrl = pageofficeCtrl.GetHtmlCode("PageOfficeCtrl1");
            ViewBag.time = DateTime.Now.ToString("yyyy年MM月dd日 dddd");
            return View();

        }

        public async Task<ActionResult> SaveDoc()
        {
            PageOfficeNetCore.FileSaver fs = new PageOfficeNetCore.FileSaver(Request, Response);
            await fs.LoadAsync();
            string webRootPath = _webHostEnvironment.WebRootPath;
            fs.SaveToFile(webRootPath + "/doc/" + fs.FileName);
            fs.Close();
            return Content("OK");
        }

        public IActionResult create()
        {

            SqliteConnection conn = new SqliteConnection(connString);
            conn.Open();
            string sql = "select Max(ID) from excel";
            SqliteCommand cmd = new SqliteCommand(sql, conn);
            cmd.CommandType = CommandType.Text;
            cmd.ExecuteNonQuery();
            cmd.CommandText = sql;
            SqliteDataReader dr = cmd.ExecuteReader();
            string newID = "1";
            if (dr.Read())
            {
                //newID = ((int)dr[0] + 1).ToString();
                newID = (int.Parse(dr[0].ToString()) + 1).ToString();
            }
            dr.Close();
            string fileName = "bbcc" + newID + ".xls";

            string FileSubject = "请输入文档主题";
            if (Request.Query["FileSubject"] != "") FileSubject = Request.Query["FileSubject"];

            string strsql = "Insert into excel(ID,FileName,Subject,SubmitTime) values(" + newID
                + ",'" + fileName + "','" + FileSubject + "','" + DateTime.Now.ToString() + "')";

            SqliteCommand cmd2 = new SqliteCommand(strsql, conn);

            cmd2.CommandType = CommandType.Text;
            cmd2.ExecuteNonQuery();
            conn.Close();
            // 复制服务器端的模板文件命名为新的文件名
            string webRootPath = _webHostEnvironment.WebRootPath;
            System.IO.File.Copy(webRootPath + "\\doc\\" + Request.Query["TemplateName"],
                webRootPath + "\\doc\\" + fileName, true);
            return Redirect("/");
        }

    }
}