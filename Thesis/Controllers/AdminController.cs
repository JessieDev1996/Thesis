using Project_Test1.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Test1.Controllers
{
    public class AdminController : Controller
    {
        // GET: Admin
        public ActionResult Index(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewBag.Page = "View";
            ConnectDB db = new ConnectDB();
            var _viewAD = "SELECT * FROM Account ;";
            _viewAD += "SELECT * FROM DefineSemester ;";
            _viewAD += "SELECT * FROM TeachersandAuthorities WHERE teacher_id = '"+ ti.user_id_title + "' ";
            DataSet ds = new DataSet();
            
            ds = db.select(_viewAD, "AD");

            ViewData["name"] = ds.Tables["AD2"].Rows[0]["name"].ToString();

            return View(ds);
        }

        public ActionResult EditProfile(TitileAll ti)
        {
            //user_id = "admin";
            ViewData["user_id"] = ti.user_id_title;
            ViewBag.Page = "Profile";
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ConnectDB db = new ConnectDB();
            Account acc = new Account();
            var _edit = "SELECT * FROM Account where user_id = '"+ ti.user_id_title + "' ;";
            _edit += "SELECT * FROM TeachersandAuthorities WHERE teacher_id = '" + ti.user_id_title + "' ";
            DataSet ds = new DataSet();
            ds = db.select(_edit, "Edit");
            /*
            acc.user_id = ds.Tables["Account"].Rows[0]["user_id"].ToString();
            acc.password = ds.Tables["Account"].Rows[0]["password"].ToString();
            acc.position = ds.Tables["Account"].Rows[0]["position"].ToString();
            acc.major = ds.Tables["Account"].Rows[0]["major"].ToString();
            acc.date_time_in = ds.Tables["Account"].Rows[0]["date_time_in"].ToString();
            acc.date_time_update = ds.Tables["Account"].Rows[0]["date_time_update"].ToString();
            */
            return View(ds);
        }
        [HttpPost]
        public ActionResult EditProfile(Account acc, TeachersandAuthorities ta, TitileAll ti)
        {
            acc.date_time_update = DateTime.Now.ToString();
            var _update = "UPDATE Account SET password = '" + acc.password + "', position = '" + acc.position + "', major = '" + acc.major + "', date_time_in = '"
                + acc.date_time_in + "', date_time_update = '" + acc.date_time_update + "' WHERE user_id = '" + acc.user_id + "' ;";
            _update += "UPDATE TeachersandAuthorities SET name = '" + ta.name + "', email = '" + ta.email + "', phone = '" + ta.phone + "', "
                + "status = '" + ta.status + "' WHERE teacher_id = '" + acc.user_id + "' ;";
            ConnectDB update = new ConnectDB();
            ti.msgAlert = update.insert_update_delete(_update);
            return RedirectToAction("EditProfile", ti); ;
        }



        //Manage Account -----------------------------------------------------
        public ActionResult AddAccount(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ViewBag.Page = "Add Account";
            
            ConnectDB db = new ConnectDB();
            var _countAccount = "select count(*) from Account ;";
            _countAccount += "select count(*) from Account where position = 'ผู้ดูแล' ;";
            _countAccount += "select count(*) from Account where position = 'เจ้าหน้าที่ฝ่ายสหกิจศึกษา' ;";
            _countAccount += "select count(*) from Account where position = 'อาจารย์ฝ่ายสหกิจศึกษา' or position = 'ผู้ช่วยอาจารย์ประจำภาควิชา' ;";
            _countAccount += "select count(*) from Account where position = 'อาจารย์นิเทศ' ";
            DataSet ds = new DataSet();
            ds = db.select(_countAccount, "Count");
            return View(ds);
        }
        [HttpPost]
        public ActionResult AddAccount(Account acc, TitileAll ti)
        {
            acc.date_time_in = DateTime.Now.ToString();
            acc.date_time_update = DateTime.Now.ToString();
            if (acc.major == null)
            {
                acc.major = "เจ้าหน้าที่ฝ่ายสหกิจศึกษา";
            }
            ConnectDB select = new ConnectDB();
            var checkAcc = "SELECT user_id FROM Account WHERE user_id = '" + acc.user_id + "'";
            DataSet ds = new DataSet();
            ds = select.select(checkAcc, "Account");
            var result = ds.Tables["Account"].Rows.Count;
            if (result == 1)
            {
                ti.msgAlert = "Already";
                return RedirectToAction("AddAccount", ti);
            }
            else
            {
                var _insert = "INSERT INTO Account VALUES ('" + acc.user_id + "', '" + acc.password + "', '" + acc.position + "', '" + acc.major +
                    "', '" + acc.date_time_in + "', '" + acc.date_time_update + "') ;";
                _insert += "INSERT INTO TeachersandAuthorities(teacher_id) VALUES ('" + acc.user_id + "') ";
                ConnectDB insert = new ConnectDB();
                string status = insert.insert_update_delete(_insert);
                ti.msgAlert = status;
                ModelState.Clear();
                return RedirectToAction("AddAccount", ti);
            }
        }


        public ActionResult DeleteAccount(string user_id_del, TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            var _delete = "DELETE FROM Account WHERE user_id = '" + user_id_del + "' ;";
            _delete += "DELETE FROM TeachersandAuthorities WHERE teacher_id = '" + user_id_del + "' ";
            string delete = db.insert_update_delete(_delete);
            ViewBag.SuccessMessage = delete;
            return RedirectToAction("Index", ti);
        }


        public ActionResult EditAccount(string user_id_edit, TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Edit Account";

            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            Account acc = new Account();
            string _edit = "SELECT * FROM Account WHERE user_id = '" + user_id_edit + "' ;";
            _edit += "SELECT * FROM TeachersandAuthorities WHERE teacher_id = '" + user_id_edit + "' ";
            ds = db.select(_edit, "Edit");
            /*acc.user_id = ds.Tables["Account"].Rows[0]["user_id"].ToString();
            acc.password = ds.Tables["Account"].Rows[0]["password"].ToString();
            acc.position = ds.Tables["Account"].Rows[0]["position"].ToString();
            acc.major = ds.Tables["Account"].Rows[0]["major"].ToString();
            acc.date_time_in = ds.Tables["Account"].Rows[0]["date_time_in"].ToString();
            acc.date_time_update = ds.Tables["Account"].Rows[0]["date_time_update"].ToString();*/
            if (ds.Tables["Edit1"].Rows[0]["name"].ToString() == "") ds.Tables["Edit1"].Rows[0]["name"] = "ไม่ระบุ";
            if (ds.Tables["Edit1"].Rows[0]["email"].ToString() == "") ds.Tables["Edit1"].Rows[0]["email"] = "ไม่ระบุ";
            if (ds.Tables["Edit1"].Rows[0]["phone"].ToString() == "") ds.Tables["Edit1"].Rows[0]["phone"] = "ไม่ระบุ";
            if (ds.Tables["Edit1"].Rows[0]["status"].ToString() == "") ds.Tables["Edit1"].Rows[0]["status"] = "ไม่ระบุ";
            return View(ds);

        }
        [HttpPost]
        public ActionResult EditAccount(Account acc, TitileAll ti)
        {
            //ViewData["user_id"] = user_id_admin;
            acc.date_time_update = DateTime.Now.ToString();
            var _update = "UPDATE Account SET password = '" + acc.password + "', position = '" + acc.position + "', major = '" + acc.major + "', date_time_in = '"
                + acc.date_time_in + "', date_time_update = '" + acc.date_time_update + "' WHERE user_id = '" + acc.user_id + "'";
            ConnectDB update = new ConnectDB();
            _ = update.insert_update_delete(_update);
            return RedirectToAction("Index", "Admin", ti);
        }



        //Manage Semester -----------------------------------------------------

        public ActionResult SetSemester(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ViewBag.Page = "Set Semester";
            ConnectDB db = new ConnectDB();
            var _viewSemester = "SELECT * FROM DefineSemester order by semester_id DESC LIMIT 1";
            DataSet ds = new DataSet();
            ds = db.select(_viewSemester, "Semester");
            if (ds.Tables["Semester"].Rows.Count == 0) ds.Tables["Semester"].Rows.Add();
            string year = ds.Tables["Semester"].Rows[0]["semester_id"].ToString();
            string[] semester = year.Split('-');
            ds.Tables["Semester"].Rows[0]["semester_id"] = semester[0];
            return View(ds);
        }
        [HttpPost]
        public ActionResult SetSemester(DefineSemester def_s, TitileAll ti)
        {
            //ViewData["user_id"] = user_id;
            string semester_id = def_s.semester_id + "-1";
            ConnectDB select = new ConnectDB();
            var checkAcc = "SELECT * FROM DefineSemester WHERE semester_id = '" + semester_id + "'";
            DataSet ds = new DataSet();
            ds = select.select(checkAcc, "DefineSemester");
            var result = ds.Tables["DefineSemester"].Rows.Count;
            if (result == 1)
            {
                ti.msgAlert = "Already";
                return RedirectToAction("SetSemester", ti);
            }
            else
            {
                var _insert = "INSERT INTO DefineSemester VALUES ('" + semester_id + "', '" + def_s.date_start1 + "', '" + def_s.date_end1 + "', " + def_s.to_go1 + ");";
                semester_id = def_s.semester_id + "-2";
                _insert += " INSERT INTO DefineSemester VALUES ('" + semester_id + "', '" + def_s.date_start2 + "', '" + def_s.date_end2 + "', " + def_s.to_go2 + ");" ;
                semester_id = def_s.semester_id + "-3";
                _insert += " INSERT INTO DefineSemester VALUES ('" + semester_id + "', '" + def_s.date_start3 + "', '" + def_s.date_end3 + "', " + def_s.to_go3 + ");";
                ConnectDB insert = new ConnectDB();
                string status = insert.insert_update_delete(_insert);
                ti.msgAlert = status;
                ModelState.Clear();
                return RedirectToAction("SetSemester", ti);
            }
        }


        public ActionResult EditSemester(string semester_id, TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Edit Semester";
            if (semester_id != null)
            {
                ConnectDB db = new ConnectDB();
                DataSet ds = new DataSet();
                //DefineSemester def_s = new DefineSemester();
                string[] year = semester_id.Split('-');
                year[0] = (int.Parse(year[0]) - 1).ToString();
                string semester = year[0] + "-" + year[1];
                string _edit = "SELECT * FROM DefineSemester WHERE semester_id = '" + semester_id + "' ;";
                _edit += "SELECT * FROM DefineSemester WHERE semester_id = '" + semester + "' ";
                ds = db.select(_edit, "Semester");
                if (ds.Tables["Semester1"].Rows.Count == 0)
                {
                    ds.Tables["Semester1"].Rows.Add();
                }
                else
                {
                    ds.Tables["Semester1"].Rows[0]["semester_id"] = year[0];
                }
                ds.Tables["Semester"].Rows[0]["semester_id"] = (int.Parse(year[0]) + 1).ToString();
                ViewData["edit"] = year[1];
                ViewData["show"] = year[1];
                if (year[1] == "3")
                {
                    ViewData["txt"] = "ฝึกงาน";
                }
                else
                {
                    ViewData["txt"] = "สหกิจ";
                }
                /*
                def_s.semester_id = ds.Tables["Semester"].Rows[0]["semester_id"].ToString();
                def_s.date_start1 = ds.Tables["Semester"].Rows[0]["date_start"].ToString();
                def_s.date_end1 = ds.Tables["Semester"].Rows[0]["date_end"].ToString();
                def_s.to_go1 = (int)ds.Tables["Semester"].Rows[0]["to_go"];
                */
                return View(ds);
            }
            else
            {
                return View();
            }

        }
        [HttpPost]
        public ActionResult EditSemester(DefineSemester def_s, TitileAll ti)
        {
            //ViewData["user_id"] = user_id;
            var _update = "UPDATE DefineSemester SET date_start = '" + def_s.date_start1 + "', date_end = '" + def_s.date_end1 + "', to_go = " + def_s.to_go1 
                + " WHERE semester_id = '" + def_s.semester_id + "'";
            ConnectDB update = new ConnectDB();
            _ = update.insert_update_delete(_update);
            return RedirectToAction("Index", ti);
        } 



        // ------------------------------------------------- เหลืออย่างเดียวของ ADMIN  ----------------------------------
        public ActionResult Report(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;

            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            //DefineSemester def_s = new DefineSemester();
            string sql = "SELECT semester_id FROM DefineSemester order by semester_id desc;";
            sql += "select distinct a.major from Student s inner join Account a on a.user_id = s.student_id";
            ds = db.select(sql, "Semester");

            return View(ds);
        }

        public ActionResult excelReport(string sem , string major)
        {
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var sql = "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out, s.sec, s.advisor from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + major + "' and(s.semester_id = '" + sem + "' or s.semester_out = '" + sem + "');";
            sql += "select company_id,name_th from Company;";

            sql += "select c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.postal_code, p.province, c.phone, c.fax, s.name_th, a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '" + major + "' and s.semester_out = '" + sem + "' order by c.name_th;";
            sql += "select  distinct c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.province, p.postal_code, c.phone, c.fax from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '" + major + "' and s.semester_out = '" + sem + "' order by c.name_th;";
            sql += "select c.name_th , ct.name, ct.position, ct.department from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id inner join Contact ct on ct.student_company_id like c.company_id where a.major = '" + major + "' and s.semester_out = '" + sem + "' order by c.name_th; ";

            sql += "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + major + "' and s.semester_out = '" + sem + "';";

            sql += "select distinct s.sec from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + major + "' and(s.semester_id = '" + sem + "' or s.semester_out = '" + sem + "');";

            ds = db.select(sql, "excel");

            DataTable dt = ds.Tables[0];
            DataTable dt_company = ds.Tables[1]; ////

            DataTable dt0 = ds.Tables[2];
            DataTable dt1 = ds.Tables[3];
            DataTable dt2 = ds.Tables[4];



            string[] split_sem = sem.Split('-');
            string year = split_sem[0];
            string term = split_sem[1];

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                application.DefaultVersion = ExcelVersion.Excel2016;

                //Create a workbook
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet worksheet = workbook.Worksheets[0];
                IWorksheet worksheet1 = workbook.Worksheets[1];

                {
                    int total = 0;
                    int count_st = 0;
                    int count = 0;
                    int count_st_of_sec = 0;

                    //Apply font setting for cells with product details
                    worksheet.Range["A1:T3000"].CellStyle.Font.FontName = "TH Sarabun New";
                    worksheet.Range["A1:T3000"].CellStyle.Font.Size = 14;

                    //Apply row height and column width to look good
                    worksheet.Range["A1"].ColumnWidth = 5;
                    worksheet.Range["B1"].ColumnWidth = 15;
                    worksheet.Range["C1"].ColumnWidth = 22;
                    worksheet.Range["D1:T1"].ColumnWidth = 3;
                    worksheet.Range["A1:A3000"].RowHeight = 20;

                    //Alignment

                    worksheet.Range["A1:A3000"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    worksheet.Range["A1:T3000"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;


                    for (int a = 0; a < ds.Tables[6].Rows.Count; a++)
                    {
                        for (int c = 0; c < dt.Rows.Count; c++)
                        {
                            if (ds.Tables[6].Rows[a][0].ToString() == dt.Rows[c][5].ToString())
                            {
                                if (dt.Rows[c][3].ToString() == dt.Rows[c][4].ToString() || dt.Rows[c][4].ToString() == sem)
                                {
                                    count_st_of_sec++;
                                }
                            }
                        }
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            //Merge 
                            if (ds.Tables[6].Rows[a][0].ToString() == dt.Rows[i][5].ToString())
                            {
                                if (count == 0) ///// เริ่ม sec ใหม่
                                {
                                    //Add a picture
                                    IPictureShape shape = worksheet.Pictures.AddPicture(total + 1, 1, Server.MapPath("~/images/logo_rmutt.jpg"));

                                    worksheet.Range["A" + (total + 1) + ":" + "A" + (total + 3)].Merge();
                                    worksheet.Range["B" + (total + 1) + ":" + "C" + (total + 1)].Merge();
                                    worksheet.Range["B" + (total + 2) + ":" + "C" + (total + 2)].Merge();
                                    worksheet.Range["B" + (total + 3) + ":" + "J" + (total + 3)].Merge();
                                    worksheet.Range["C" + (total + 4) + ":" + "J" + (total + 4)].Merge();

                                    worksheet.Range["K" + (total + 1) + ":" + "T" + (total + 1)].Merge();
                                    worksheet.Range["K" + (total + 2) + ":" + "T" + (total + 2)].Merge();
                                    worksheet.Range["K" + (total + 3) + ":" + "T" + (total + 3)].Merge();
                                    worksheet.Range["K" + (total + 4) + ":" + "T" + (total + 4)].Merge();

                                    worksheet.Range["B" + (total + 5) + ":" + "D" + (total + 5)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    worksheet.Range["B" + (total + 1) + ":" + "B" + (total + 4)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                    worksheet.Range["K" + (total + 1) + ":" + "K" + (total + 4)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignLeft;
                                    worksheet.Range["C" + (total + 4)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 5) + ":" + "T" + (total + 5)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                                    worksheet.Range["D" + (total + 5) + ":" + "T" + (total + 5)].Merge();

                                    worksheet.Range["A" + (total + 5) + ":" + "D" + (total + 5)].CellStyle.Font.Bold = true;
                                    worksheet.Range["B" + (total + 1) + ":" + "B" + (total + 4)].CellStyle.Font.Bold = true;
                                    worksheet.Range["K" + (total + 1) + ":" + "K" + (total + 4)].CellStyle.Font.Bold = true;
                                    worksheet.Range["C" + (total + 4)].CellStyle.Font.Bold = true;

                                    worksheet.Range["B" + (total + 1)].Text = "มหาวิทยาลัยเทคโนโลยีราชมงคลธัญบุรี";
                                    worksheet.Range["B" + (total + 2)].Text = "คณะวิศวกรรมศาสตร์";
                                    worksheet.Range["B" + (total + 3)].Text = "หลักสูตร 2004060206  :  " + major;
                                    worksheet.Range["B" + (total + 4)].Text = "ปริญญาตรี";

                                    worksheet.Range["K" + (total + 1)].Text = "รายชื่อนักศึกษา";
                                    worksheet.Range["K" + (total + 2)].Text = "ภาคการศึกษาที่ " + term + "/" + year;
                                    worksheet.Range["K" + (total + 3)].Text = "กลุ่ม " + ds.Tables[6].Rows[a][0].ToString();
                                    worksheet.Range["K" + (total + 4)].Text = "ภาคปกติ  จำนวน " + count_st_of_sec.ToString() + " คน";

                                    worksheet.Range["A" + (total + 5)].Text = "ลำดับ";
                                    worksheet.Range["B" + (total + 5)].Text = "รหัสนักศึกษา";
                                    worksheet.Range["C" + (total + 5)].Text = "ชื่อ - สกุล";
                                    worksheet.Range["D" + (total + 5)].Text = "สถานประกอบการ";

                                    worksheet.Range["C" + (total + 4)].Text = "อ. ที่ปรึกษา " + dt.Rows[i][6].ToString();
                                }

                                //Apply borders

                                if (dt.Rows[i][3].ToString() == dt.Rows[i][4].ToString() || dt.Rows[i][4].ToString() == sem)
                                {
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                                    worksheet.Range["D" + (total + 6) + ":" + "T" + (total + 6)].Merge();

                                    worksheet.Range["A" + (total + 6)].Text = (count + 1).ToString();
                                    worksheet.Range["B" + (total + 6)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    worksheet.Range["B" + (total + 6)].Text = dt.Rows[i][0].ToString();
                                    worksheet.Range["C" + (total + 6)].Text = dt.Rows[i][1].ToString();

                                    for (int c = 0; c < dt_company.Rows.Count; c++)
                                    {
                                        if (dt.Rows[i][2].ToString() == dt_company.Rows[c][0].ToString())
                                        {
                                            worksheet.Range["D" + (total + 6)].Text = dt_company.Rows[c][1].ToString(); ////
                                            count_st++; //////////////////
                                        }
                                    }

                                    total++;
                                }

                                if (dt.Rows[i][3].ToString() != dt.Rows[i][4].ToString() && dt.Rows[i][3].ToString() == sem)
                                {
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                                    worksheet.Range["A" + (total + 6) + ":" + "T" + (total + 6)].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                                    worksheet.Range["D" + (total + 6) + ":" + "T" + (total + 6)].Merge();

                                    worksheet.Range["A" + (total + 6)].Text = (count + 1).ToString();
                                    worksheet.Range["B" + (total + 6)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    worksheet.Range["B" + (total + 6)].Text = dt.Rows[i][0].ToString();
                                    worksheet.Range["C" + (total + 6)].Text = dt.Rows[i][1].ToString();
                                    worksheet.Range["D" + (total + 6)].Text = "ไม่ออกฝึก";

                                    worksheet.Range["A" + (total + 6) + ":" + "D" + (total + 6)].CellStyle.Font.Color = ExcelKnownColors.Red;

                                    total++;
                                }

                                count++;
                            }
                        }
                        total += 10;
                        count = 0;  /////////////////////////////
                        count_st_of_sec = 0;
                    }
                }

                {
                    int total = 3;
                    int nub = 0;
                    int total_student = 0;
                    int count_student = 0;

                    worksheet1.Range["A1:T3000"].CellStyle.Font.FontName = "TH Sarabun New";
                    worksheet1.Range["A1:T3000"].CellStyle.Font.Size = 14;

                    //Apply row height and column width to look good
                    worksheet1.Range["A1"].ColumnWidth = 10;
                    worksheet1.Range["B1"].ColumnWidth = 35;
                    worksheet1.Range["C1"].ColumnWidth = 35;
                    worksheet1.Range["D1:F1"].ColumnWidth = 25;
                    worksheet1.Range["A1:A3000"].RowHeight = 25;

                    //Alignment

                    worksheet1.Range["A1:A3000"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    worksheet1.Range["A2:F2"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    worksheet1.Range["F2:F3000"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                    worksheet1.Range["A1:F3000"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;

                    worksheet1.Range["A1:F1"].Merge();
                    worksheet1.Range["A1"].CellStyle.Font.Bold = true;
                    worksheet1.Range["A2:F2"].CellStyle.Font.Bold = true;

                    worksheet1.Range["A2"].Text = "ลำดับที่";
                    worksheet1.Range["B2"].Text = "ชื่อหน่วยงาน";
                    worksheet1.Range["C2"].Text = "ที่อยู่";
                    worksheet1.Range["D2"].Text = "ผู้ติดต่อ";
                    worksheet1.Range["E2"].Text = "รายชื่อนักศึกษา";
                    worksheet1.Range["F2"].Text = "สาขาวิชา";

                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeTop].Color = ExcelKnownColors.Black;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                    worksheet1.Range["A2:F2"].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                    for (int i = 0; i < dt1.Rows.Count; i++)
                    {
                        int check_moo;
                        bool result = Int32.TryParse(dt1.Rows[i][4].ToString(), out check_moo);

                        worksheet1.Range["A" + total].Text = (i + 1).ToString();
                        worksheet1.Range["B" + total].Text = dt1.Rows[i][0].ToString();

                        ///////////////////////////////////// รายชื่อผู้ติดต่อ
                        for (int supervise = 0; supervise < dt2.Rows.Count; supervise++)
                        {
                            if (dt1.Rows[i][0].ToString() == dt2.Rows[supervise][0].ToString())
                            {
                                worksheet1.Range["D" + (total)].Text = dt2.Rows[supervise][1].ToString();
                                worksheet1.Range["D" + (total + 1)].Text = dt2.Rows[supervise][2].ToString();
                                //total++;
                            }
                        }
                        ////////////////////////////////////

                        if (result == false)
                        {
                            if (dt1.Rows[i][4].ToString() == null || dt1.Rows[i][4].ToString() == "")
                            {
                                worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
                            }
                            else
                            {
                                worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " อาคาร " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
                            }

                        }
                        else
                        {
                            if (dt1.Rows[i][4].ToString() == null || dt1.Rows[i][4].ToString() == "")
                            {
                                worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
                            }
                            else
                            {
                                worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " หมู่ " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
                            }
                        }

                        //worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
                        worksheet1.Range["C" + (total + 1)].Text = "ตำบล " + dt1.Rows[i][6].ToString() + " อำเภอ " + dt1.Rows[i][7].ToString();
                        worksheet1.Range["C" + (total + 2)].Text = "จังหวัด " + dt1.Rows[i][8].ToString() + " " + dt1.Rows[i][9].ToString();
                        worksheet1.Range["C" + (total + 3)].Text = "โทร : " + dt1.Rows[i][10].ToString();
                        worksheet1.Range["C" + (total + 4)].Text = "แฟกซ์ : " + dt1.Rows[i][11].ToString();
                        //worksheet.Range["D" + total].Text = "ผู้ติดต่อ";

                        for (int j = 0; j < dt0.Rows.Count; j++)
                        {
                            if (dt1.Rows[i][0].ToString() == dt0.Rows[j][0].ToString())
                            {
                                worksheet1.Range["E" + total].Text = (count_student + 1).ToString() + ". " + dt0.Rows[j][12].ToString();
                                worksheet1.Range["F" + total].Text = dt0.Rows[j][13].ToString();

                                worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                                worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                                worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                                worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                                total++;
                                total_student++;
                                count_student++;
                                nub++;
                            }
                        }



                        for (int a = 0; nub < 5; a++) /// ถ้าเด็กน้อยกว่า 4 คน
                        {
                            worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                            worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                            worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                            worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;

                            nub++;
                            total++;
                        }

                        nub = 0;
                        count_student = 0;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeBottom].Color = ExcelKnownColors.Black;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeRight].Color = ExcelKnownColors.Black;
                        worksheet1.Range["A" + total + ":" + "F" + total].CellStyle.Borders[ExcelBordersIndex.EdgeLeft].Color = ExcelKnownColors.Black;
                        total++;
                    }

                    worksheet1.Range["A1"].Text = "ข้อมูลสถานประกอบการที่ส่งนักศึกษาสหกิจศึกษา ภาคเรียนที่ " + term + "/" + year + " เข้าฝึกประสบการณ์วิชาชีพ ภาควิชา" + major + " จำนวน " + total_student + " คน";

                }

                worksheet.Name = "รายชื่อนักศึกษา " + sem;
                worksheet1.Name = "รายชื่อบริษัท " + sem;
                //Save the workbook to disk in xlsx format
                workbook.SaveAs("รายงานการออกฝึกงานภาควิชา" + major + " ปี " + sem + ".xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
            }

            return null;
        }
    }
}