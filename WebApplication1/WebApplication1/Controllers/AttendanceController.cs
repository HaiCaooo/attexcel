using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using WebApplication1.Models;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace WebApplication1.Controllers
{
    public class AttendanceController : Controller
    {
        cap21t41Entities db = new cap21t41Entities();
        // GET: Attendance
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ImportFile()
        {
            return View();
        }
        public ActionResult Success()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ImportFile(HttpPostedFileBase file)
        {
            if (file.ContentLength ==0 || file == null)
            {
                ViewBag.Error = "Please select excel file";
                return View();
            }
            else
            {
                if (file.FileName.EndsWith("xls") || file.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Files/"  + file.FileName);
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                    }
                    file.SaveAs(path);

                    //read excel 
                    Excel.Application application = new Excel.Application();
                    Excel.Workbook workbook = application.Workbooks.Open(path);
                    Excel.Worksheet worksheet = workbook.ActiveSheet;
                    Excel.Range range = worksheet.UsedRange;
                    var a = range.Rows.Count;
                    List<Attendance> lsatd = new List<Attendance>();

                    for (int i = 4; i <= range.Rows.Count; i++)
                    {
                        Attendance atd = new Attendance();
                        //var mid = ((Excel.Range)range.Cells[i, 1]).Text;
                        var memid = Convert.ToInt32(((Excel.Range)range.Cells[i, 1]).Text);
                        atd.MemberID = Convert.ToInt32(memid);
                        atd.SessionID = Convert.ToInt32(((Excel.Range)range.Cells[i, 2]).Text);
                        atd.Status = Convert.ToByte(((Excel.Range)range.Cells[i, 3]).Text);
                        atd.Note = Convert.ToString(((Excel.Range)range.Cells[i, 4]).Text);
                        lsatd.Add(atd);                
                    }
                    ViewBag.lsatd = lsatd;
                    
                    foreach (var item in lsatd)
                    {
                        Attendance att = new Attendance();
                        att.MemberID = item.MemberID;
                        att.SessionID = item.SessionID;
                        att.Status = item.Status;
                        att.Note = item.Note;
                        db.Attendances.Add(att);
                    }
                    db.SaveChanges();
                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "Please select correct excel file";
                    return View();
                }
                
            }
            
        }
         

        //[HttpPost]
        //public ActionResult ImportFile(HttpPostedFileBase file)
        //{
        //    DataSet ds = new DataSet();
        //    if (Request.Files["file"].ContentLength > 0)
        //    {
        //        string fileExtension = System.IO.Path.GetExtension(Request.Files["file"].FileName);

        //        if (fileExtension == ".xls" || fileExtension == ".xlsx")
        //        {
        //            string fileLocation = Server.MapPath("~/Files/") + Request.Files["file"].FileName;
        //            if (System.IO.File.Exists(fileLocation))
        //            {

        //                System.IO.File.Delete(fileLocation);
        //            }
        //            Request.Files["file"].SaveAs(fileLocation);
        //            string excelConnectionString = string.Empty;
        //            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
        //            fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
        //            //connection String for xls file format.
        //            if (fileExtension == ".xls")
        //            {
        //                excelConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" +
        //                fileLocation + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
        //            }
        //            //connection String for xlsx file format.
        //            else if (fileExtension == ".xlsx")
        //            {
        //                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
        //                fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
        //            }

        //            //Create Connection to Excel work book and add oledb namespace
        //            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
        //            excelConnection.Open();
        //            DataTable dt = new DataTable();

        //            dt = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
        //            if (dt == null)
        //            {
        //                return null;
        //            }

        //            String[] excelSheets = new String[dt.Rows.Count];
        //            int t = 0;
        //            //excel data saves in temp file here.
        //            foreach (DataRow row in dt.Rows)
        //            {
        //                excelSheets[t] = row["TABLE_NAME"].ToString();
        //                t++;
        //            }
        //            OleDbConnection excelConnection1 = new OleDbConnection(excelConnectionString);


        //            string query = string.Format("Select * from [{0}]", excelSheets[0]);
        //            using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, excelConnection1))
        //            {
        //                dataAdapter.Fill(ds);
        //            }
        //        }               

        //        for (int i = 3; i < ds.Tables[0].Rows.Count; i++)
        //        {
        //            var atd = new Attendance();
        //            var memid = (int)ds.Tables[0].Rows[i][0];
        //            //atd.MemberID. = 
        //            string s1 = (string)ds.Tables[0].Rows[i][3];


        //            atd.SessionID = DateTime.Now.Day;
        //            db.Attendances.Add(atd);
        //            db.SaveChanges();
        //        }
        //    }
        //    return RedirectToAction("Index", "Home");
        //}
    }
}
