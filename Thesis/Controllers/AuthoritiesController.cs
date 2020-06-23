using iTextSharp.text;
using iTextSharp.text.pdf;
using Project_Test1.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Test1.Controllers
{
    public class AuthoritiesController : Controller
    {

        // GET: Authorities
        public ActionResult Index(TitileAll ti)
        {
            //ti.user_id_title = "kittisak";
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "View";
            ConnectDB db = new ConnectDB(); 
            var _viewCompany = "SELECT company_id, name_en, name_th, website, blacklist FROM Company ;";
            _viewCompany += "SELECT * FROM TeachersandAuthorities WHERE teacher_id = '" + ti.user_id_title + "' ";
            DataSet ds = new DataSet();
            ds = db.select(_viewCompany, "Company");
            ViewData["name"] = ds.Tables["Company1"].Rows[0]["name"].ToString();
            return View(ds);
        }



        public ActionResult No02()
        {
            string ReportURL = Server.MapPath("~/prototypeNo/no02.pdf");
            byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
            return File(FileBytes, "application/pdf");
        }



        public ActionResult EditProfile(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewBag.Page = "Profile";
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ConnectDB db = new ConnectDB();
            var _edit = "SELECT * FROM Account where user_id = '" + ti.user_id_title + "' ;";
            _edit += "SELECT * FROM TeachersandAuthorities WHERE teacher_id = '" + ti.user_id_title + "' ";
            DataSet ds = new DataSet();
            ds = db.select(_edit, "Edit");
            return View(ds);
        }
        [HttpPost]
        public ActionResult EditProfile(Account acc, TeachersandAuthorities ta, TitileAll ti)
        {
            acc.date_time_update = DateTime.Now.ToString();
            var _update = "UPDATE Account SET password = '" + acc.password + "', position = '" + acc.position + "', major = '" + acc.major + "', date_time_in = '"
                + acc.date_time_in + "', date_time_update = '" + acc.date_time_update + "' WHERE user_id = '" + acc.user_id + "' ;";
            _update += "UPDATE TeachersandAuthorities SET name = '" + ta.name + "', email = '" + ta.email + "', phone = '" + ta.phone + "', "
                + "status = '" + ta.status + "' WHERE teacher_id = '" + acc.user_id + "' ";
            ConnectDB update = new ConnectDB();
            ti.msgAlert = update.insert_update_delete(_update);
            return RedirectToAction("EditProfile", ti); ;
        }



        //Manage Company -----------------------------------------------------------------
        public ActionResult AddCompany(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ViewBag.Page = "Add Company";
            
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            List<string> ListItems = new List<string>();
            var _Address = "SELECT DISTINCT province FROM PrototypeAddress where not prototype_id = 0 order by province ;";
            ds = db.select(_Address, "PrototypeAddress");
            int _count = ds.Tables["PrototypeAddress"].Rows.Count;
            for (int i = 0; i < _count; i++)
            {
                ListItems.Add(ds.Tables["PrototypeAddress"].Rows[i]["province"].ToString());
            }
            ViewBag.province = ListItems;

            return View();
        }

        [HttpPost]
        public ActionResult AddCompany(Company comp, TitileAll ti)
        {
            ConnectDB select = new ConnectDB();
            var run_id_checkCom = "SELECT MAX(company_id) FROM Company ;SELECT MAX(address_id) FROM Address ;SELECT MAX(contact_id) FROM Contact ;SELECT MAX(rs_id) FROM ReceiveSemester ;";
            run_id_checkCom += "SELECT name_en, name_th, website FROM Company WHERE name_en = '" + comp.name_en[0] + "' or name_th = '" + comp.name_th[0] + "';";
            DataSet ds = new DataSet();
            ds = select.select(run_id_checkCom, "Company_Address_Contact");
            var _chackCom = ds.Tables[4].Rows.Count;

            if (_chackCom == 1)
            {
                ti.msgAlert = "Already";
                //ModelState.Clear();
                return RedirectToAction("AddCompany", ti);
            }
            else
            {
                var _countCom = ds.Tables[0].Rows[0][0].ToString();
                var _countAdd = ds.Tables[1].Rows[0][0].ToString();
                var _countCon = ds.Tables[2].Rows[0][0].ToString();
                var _countRec = ds.Tables[3].Rows[0][0].ToString();

                comp.company_id = new int[1];
                //comp.name_en = new string[1];
                //comp.name_th = new string[1];
                //comp.blacklist = new string[1];

                if (_countCom == "") comp.company_id[0] = 1;
                else comp.company_id[0] = Int32.Parse(_countCom) + 1;
                if (_countAdd == "") comp.address_id = 1;
                else comp.address_id = Int32.Parse(_countAdd) + 1;
                if (_countCon == "") comp.contact_id = 1;
                else comp.contact_id = Int32.Parse(_countCon) + 1;
                if (_countRec == "") comp.rs_id = 1;
                else comp.rs_id = Int32.Parse(_countRec) + 1;

                comp.start_time = DateTime.Now.ToString();
                comp.update_time = DateTime.Now.ToString();
                comp.student_company_id = comp.company_id[0].ToString();
                var _insert = "INSERT INTO Address VALUES (" + comp.address_id + ", " + comp.prototype_id + ", " + comp.company_id[0] + ", 'สถานประกอบการ', '" + comp.number + "', '"
                    + comp.building + "', '" + comp.floor + "', '" + comp.lane + "', '" + comp.road + "' , '" + comp.pX + "' , '" + comp.pY + "') ;";
                _insert += "INSERT INTO Company VALUES (" + comp.company_id[0] + ", " + comp.address_id + ", '" + comp.name_en[0] + "', '" + comp.name_th[0] + "', 'No', '"
                    + comp.phone + "', '" + comp.fax + "', '" + comp.email + "', '" + comp.website + "', '" + comp.business_type + "', '" + comp.all_employee +
                    "', '" + comp.job_request + "', '" + comp.job_time + "', '" + comp.job_day + "', '" + comp.pay + "', '" + comp.accommodation + "', '"
                    + comp.charge_accommodation + "', '" + comp.bus + "', '" + comp.charge_bus + "', '" + comp.welfare + "', '" + comp.exam + "', '"
                    + comp.start_time + "', '" + comp.update_time + "') ;";
                if (comp.con_number == 2 && comp.name != null)
                {
                    _insert += "INSERT INTO Contact VALUES (" + comp.contact_id + ", '" + comp.student_company_id + "', 1, '" + comp.name_s + "', '" + comp.position_s +
                        "', '" + comp.department_s + "', '-', '-', '-') ;";
                    comp.contact_id += 1;
                    _insert += "INSERT INTO Contact VALUES (" + comp.contact_id + ", '" + comp.student_company_id + "', 2, '" + comp.name + "', '" + comp.position +
                        "', '" + comp.department + "', '"
                        + comp.con_phone + "', '" + comp.con_fax + "', '" + comp.con_email + "') ;";
                }
                else
                {
                    _insert += "INSERT INTO Contact VALUES (" + comp.contact_id + ", '" + comp.student_company_id + "', 1, '" + comp.name_s + "', '" + comp.position_s +
                        "', '" + comp.department_s + "', '-', '-', '-') ;";
                }
                if (comp.semester_id1 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-1', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }
                if (comp.semester_id2 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-2', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }
                if (comp.semester_id3 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-3', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }
                ConnectDB insert = new ConnectDB();
                string status = insert.insert_update_delete(_insert);
                ti.msgAlert = status;
                //ViewBag.SuccessMessage = status;
                //ModelState.Clear();
                return RedirectToAction("AddCompany", ti);
            }
        }



        public ActionResult EditCompany(Nullable<int> company_id, TitileAll ti)
        {
            /*ti.name_title = "นายกิตติศักดิ์ โคมเดือน";
            ti.user_id_title = "1160304620454";
            company_id = 2;*/
            
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Edit Company";
            
            Company comp = new Company();

                ConnectDB db = new ConnectDB();
                DataSet ds = new DataSet();
                string _edit = "SELECT * FROM Company c, Contact co, Address a, PrototypeAddress p WHERE c.company_id = " + company_id + " AND co.student_company_id = '" + company_id 
                    + "' AND a.student_company_id = '" + company_id + "' AND p.prototype_id = a.prototype_id ;";
                _edit += "SELECT DISTINCT province FROM PrototypeAddress where not prototype_id = 0 order by province ;";
                _edit += "select distinct area from PrototypeAddress where province = (select province from PrototypeAddress where prototype_id = (SELECT p.prototype_id FROM Company c,"
                    + " Address a, PrototypeAddress p WHERE c.company_id = " + company_id + " AND a.student_company_id = '" + company_id + "' AND p.prototype_id = a.prototype_id)) order by area ;";
                _edit += "select distinct sub_area from PrototypeAddress where area = (select area from PrototypeAddress where prototype_id = (SELECT p.prototype_id FROM Company c,"
                    + " Address a, PrototypeAddress p WHERE c.company_id = " + company_id + " AND a.student_company_id = '" + company_id + "' AND p.prototype_id = a.prototype_id)) order by sub_area ;";
                _edit += "SELECT p.prototype_id FROM Company c, Address a, PrototypeAddress p WHERE c.company_id = " + company_id + " AND a.student_company_id = '" + company_id + "' AND p.prototype_id = a.prototype_id;";
                ds = db.select(_edit, "Company");

            ViewData["con_number"] = ds.Tables["Company"].Rows.Count;

                if (ds.Tables["Company"].Rows.Count == 1)
                {
                    ds.Tables["Company"].Rows.Add();
                }

                List<string> ListProvince = new List<string>();
                int _countProvince = ds.Tables["Company1"].Rows.Count;
                for (int i = 0; i < _countProvince; i++)
                {
                    ListProvince.Add(ds.Tables["Company1"].Rows[i]["province"].ToString());
                }
                ViewBag.province = ListProvince;

                List<string> ListArea = new List<string>();
                int _countArea = ds.Tables["Company2"].Rows.Count;
                for (int i = 0; i < _countArea; i++)
                {
                    ListArea.Add(ds.Tables["Company2"].Rows[i]["area"].ToString());
                }

                ViewBag.area = ListArea;

                List<string> ListSubArea = new List<string>();
                int _countSubArea = ds.Tables["Company3"].Rows.Count;
                for (int i = 0; i < _countSubArea; i++)
                {
                    ListSubArea.Add(ds.Tables["Company3"].Rows[i]["sub_area"].ToString());
                }

                ViewBag.sub_area = ListSubArea;

                ViewData["prototype_id"] = ds.Tables["Company4"].Rows[0]["prototype_id"].ToString();

            return View(ds);

        }
        [HttpPost]
        public ActionResult EditCompany(Company comp, TitileAll ti)
        {
            comp.update_time = DateTime.Now.ToString();

            var _update = "UPDATE Company SET name_en = '" + comp.name_en[0] + "', name_th = '" + comp.name_th[0] + "', blacklist = '" + comp.blacklist[0] +
                "', phone = '" + comp.phone + "', fax = '" + comp.fax + "', email = '" + comp.email + "', website = '" + comp.website +
                "', business_type = '" + comp.business_type + "', all_employee = '" + comp.all_employee + "', job_request = '" + comp.job_request +
                "', job_time = '" + comp.job_time + "', job_day = '" + comp.job_day + "', pay = '" + comp.pay + "', accommodation = '" + comp.accommodation +
                "', charge_acc = '" + comp.charge_accommodation + "', bus = '" + comp.bus + "', charge_bus = '" + comp.charge_bus + "', welfare = '"
                + comp.welfare + "', exam = '" + comp.exam + "', update_time = '" + comp.update_time + "' WHERE company_id = " + comp.company_id[0] +";";

            _update += " UPDATE Address SET prototype_id = "+ comp.prototype_id +", number = '"+ comp.number + "', building = '" + comp.building + "', floor = '" + comp.floor +
                    "', lane = '" + comp.lane + "', road = '" + comp.road + "', px = '"+comp.pX+"', py = '"+comp.pY+"' WHERE address_id = " + comp.address_id + ";";

            _update += " UPDATE Contact SET name = '" + comp.name_s + "', position = '" + comp.position_s + "', department = '" + comp.department_s +
                "' WHERE student_company_id = '" + comp.company_id[0] + "' AND number = 1 ;";
            if (comp.rs_id == 2 && (comp.con_number == 1 || comp.name == null))
            {
                _update += "DELETE FROM Contact WHERE student_company_id = '" + comp.company_id[0] + "' AND number = 2 ;";
            }
            else if (comp.rs_id == 1 && comp.name != null)
            {
                ConnectDB select = new ConnectDB();
                var run_id_checkCon = "SELECT MAX(contact_id) FROM Contact";
                DataSet ds = new DataSet();
                ds = select.select(run_id_checkCon, "Company_Address_Contact");
                comp.contact_id = int.Parse(ds.Tables[0].Rows[0][0].ToString()) + 1;
                _update += "INSERT INTO Contact VALUES (" + comp.contact_id + ", '" + comp.company_id[0] + "', 2, '" + comp.name + "', '" + comp.position + "', '"
                    + comp.department + "', '" + comp.con_phone + "', '" + comp.con_fax + "', '" + comp.con_email + "') ;";
            }
            else if (comp.rs_id == 2 && comp.name != null)
            {
                _update += "UPDATE Contact SET name = '" + comp.name + "', position = '" + comp.position + "', department = '" + comp.department +
                    "', phone = '" + comp.con_phone + "', fax = '" + comp.con_fax + "', email = '" + comp.con_email +
                    "' WHERE student_company_id = '" + comp.company_id[0] + "' AND number = 2 ;";
            }

            ConnectDB update = new ConnectDB();
            _ = update.insert_update_delete(_update);
            return RedirectToAction("Index", ti);
        }



        public ActionResult DeleteCompany(int company_id, TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            var _delete = "DELETE FROM Company WHERE company_id = '" + company_id + "' ;DELETE FROM ReceiveSemester WHERE company_id = '" + company_id +
                   "'; DELETE FROM Address WHERE student_company_id = '" + company_id + "'; DELETE FROM Contact WHERE student_company_id = '" + company_id + "' ;";
            string delete = db.insert_update_delete(_delete);
            ViewBag.SuccessMessage = delete;
            return RedirectToAction("Index", ti);
        }



        // Manage Require Student ---------------------------------------------------------------
        public ActionResult AddRequire_s(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewData["msgAlert"] = ti.msgAlert;
            ViewBag.Page = "Add Require Student";
            ConnectDB db = new ConnectDB();
            var _viewCompany = "SELECT company_id, name_en, name_th, blacklist FROM Company";
            DataSet ds = new DataSet();
            ds = db.select(_viewCompany, "Company");

            Company comp = new Company();
            int _count = ds.Tables["Company"].Rows.Count;
            
            comp.blacklist = new string[_count];
            int bla = 0;
            for (int i=0; i<_count; i++)
            {
                comp.blacklist[i] = ds.Tables["Company"].Rows[i]["blacklist"].ToString();
                if (comp.blacklist[i] == "Yes") bla ++;
            }

            comp.company_id = new int[_count - bla];
            comp.name_en = new string[_count - bla];
            comp.name_th = new string[_count - bla];

            bla = 0;

            for (int i=0; i<_count; i++)
            {
                comp.blacklist[i] = ds.Tables["Company"].Rows[i]["blacklist"].ToString();
                if (comp.blacklist[i] == "Yes") continue;
                comp.company_id[bla] = int.Parse(ds.Tables["Company"].Rows[i]["company_id"].ToString());
                comp.name_en[bla] = ds.Tables["Company"].Rows[i]["name_en"].ToString();
                comp.name_th[bla] = ds.Tables["Company"].Rows[i]["name_th"].ToString();

                if (comp.name_th[bla].Length < 10)
                {
                    comp.name_th[bla] = comp.name_en[i];
                }
                bla++;
            }
            /*
            if (status == "มีข้อมูลอยู่แล้ว")
            {
                ViewBag.DuplicateMessage = status;
            }
            else if (status == "บันทึกข้อมูลสำเร็จ")
            {
                ViewBag.SuccessMessage = status;
            }*/

            return View(comp);
        }
        [HttpPost]
        public ActionResult AddRequire_s(Company comp, TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            var _check = "SELECT MAX(rs_id) FROM ReceiveSemester ;";
            
            for (int i=0; i< comp.major.Length; i++)
            {
                if (comp.semester_id1 == 1)
                {
                    _check += "SELECT semester_id FROM ReceiveSemester WHERE semester_id = '" + comp.year + "-1' AND company_id = " + comp.company_id[0] + " AND major = '" + comp.major[i] + "'; ";
                    comp.rs_id++;
                }
                if (comp.semester_id2 == 1)
                {
                    _check += "SELECT semester_id FROM ReceiveSemester WHERE semester_id = '" + comp.year + "-2' AND company_id = " + comp.company_id[0] + " AND major = '" + comp.major[i] + "'; ";
                    comp.rs_id++;
                }
                if (comp.semester_id3 == 1)
                {
                    _check += "SELECT semester_id FROM ReceiveSemester WHERE semester_id = '" + comp.year + "-3' AND company_id = " + comp.company_id[0] + " AND major = '" + comp.major[i] + "' ;";
                    comp.rs_id++;
                }
            }
            
            int[] _RepeatChack = new int[comp.rs_id];
            string status = null;
            DataSet ds = new DataSet();
            ds = db.select(_check, "ReceiveSemester");

            for (int i = 0; i < _RepeatChack.Length; i++)
            {
                _RepeatChack[i] = ds.Tables[i+1].Rows.Count;
            }


            for (int i = 0; i < _RepeatChack.Length; i++)
            {
                if (_RepeatChack[i] != 0)
                {
                    status = "Already";
                    break;
                }
            }

            if (status == "Already")
            {
                // มีหน้าที่แค่เช็ค
            }
            else
            {
                var _countRS = ds.Tables["ReceiveSemester"].Rows[0][0].ToString();

                if (_countRS == "") comp.rs_id = 1;
                else comp.rs_id = Int32.Parse(_countRS) + 1;

                string _insert = null;

                if (comp.semester_id1 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-1', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }
                if (comp.semester_id2 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-2', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }
                if (comp.semester_id3 == 1 && comp.major != null)
                {
                    for (int i = 0; i < comp.major.Length; i++)
                    {
                        _insert += "INSERT INTO ReceiveSemester VALUES (" + comp.rs_id + ", '" + comp.year + "-3', " + comp.company_id[0] + ", '" + comp.major[i] +
                            "', '" + comp.position_of_student[i] + "', '" + comp.job_description[i] + "', " + comp.require[i] + ") ;";
                        comp.rs_id += 1;
                    }
                }

                status = db.insert_update_delete(_insert);
            }

            ti.msgAlert = status;
            //ModelState.Clear();
            return RedirectToAction("AddRequire_s", ti);
        }



        public ActionResult ViewRequire_s(Nullable<int> company, string name, string year, TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "View Require Student";
            ConnectDB db = new ConnectDB();
            var _viewCompany = "SELECT company_id, name_en, name_th, blacklist FROM Company ;";
            DataSet ds = new DataSet();
            Company comp = new Company();
            string[] _year;

            if (company != null)
            {
                _year = year.Split(null);
                year = _year[1];
                comp.name = name;
                _viewCompany += "SELECT * FROM ReceiveSemester WHERE company_id = " + company + " AND semester_id LIKE '" + year + "%' ORDER BY semester_id ASC ;";
            }
            else comp.name = "Please Select";

            ds = db.select(_viewCompany, "Company");

            int _count = ds.Tables["Company"].Rows.Count;

            comp.blacklist = new string[_count];
            int bla = 0;
            for (int i = 0; i < _count; i++)
            {
                comp.blacklist[i] = ds.Tables["Company"].Rows[i]["blacklist"].ToString();
                if (comp.blacklist[i] == "Yes") bla++;
            }

            comp.company_id = new int[_count - bla];
            comp.name_en = new string[_count - bla];
            comp.name_th = new string[_count - bla];

            bla = 0;

            for (int i = 0; i < _count; i++)
            {
                comp.blacklist[i] = ds.Tables["Company"].Rows[i]["blacklist"].ToString();
                if (comp.blacklist[i] == "Yes") continue;
                comp.company_id[bla] = int.Parse(ds.Tables["Company"].Rows[i]["company_id"].ToString());
                comp.name_en[bla] = ds.Tables["Company"].Rows[i]["name_en"].ToString();
                comp.name_th[bla] = ds.Tables["Company"].Rows[i]["name_th"].ToString();

                if (comp.name_th[bla].Length < 10)
                {
                    comp.name_th[bla] = comp.name_en[i];
                }
                bla++;
            }

            if (company != null)
            {
                comp.rs_id = Int32.Parse(year);
                if (ds.Tables[1].Rows.Count != 0)
                {
                    comp.major = new string[ds.Tables[1].Rows.Count];
                    comp.require = new int[ds.Tables[1].Rows.Count];
                    comp.term = new string[ds.Tables[1].Rows.Count];
                    for (int i = 0; i < ds.Tables[1].Rows.Count; i++)
                    {
                        if (ds.Tables[1].Rows[i]["semester_id"].ToString() == year + "-1")
                        {
                            comp.term[i] = "1";
                        }
                        else if (ds.Tables[1].Rows[i]["semester_id"].ToString() == year + "-2")
                        {
                            comp.term[i] = "2";
                        }
                        else if (ds.Tables[1].Rows[i]["semester_id"].ToString() == year + "-3")
                        {
                            comp.term[i] = "3";
                        }
                        comp.major[i] = ds.Tables[1].Rows[i]["major"].ToString();
                        comp.require[i] = int.Parse(ds.Tables[1].Rows[i]["require"].ToString());
                    }
                }
            }
            else
            {
                comp.rs_id = DateTime.Now.Year + 543;
            }

            return View(comp);
        }





        // Confirm Student ---------------------------------------------------------------
        public ActionResult ConfirmStudent(Nullable<int> company_id, string company_name, TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Confirm Company";
            ConnectDB db = new ConnectDB();
            var confirm = "select * from Company where company_id in (select company_id from Student where confirm_status = 'รอการตอบรับ') ;";
            //confirm += "select company_id, student_id, name_th, confirm_status from Student where confirm_status = 'รอการตอบรับ' ";
            DataSet ds = new DataSet();

            if(company_id == null)
            {
                ViewData["company_name"] = "กรุณาเลือกสถานประกอบการ";
            }
            else
            {
                ViewData["company_name"] = company_name;
                confirm += "select s.company_id, s.student_id, s.name_th, a.major, s.confirm_status from Student s, Account a where confirm_status = 'รอการตอบรับ'"
                    + " and a.user_id = s.student_id and s.company_id = "+ company_id +" order by a.major, s.student_id ;";
            }

            ds = db.select(confirm, "Company_student");

            return View(ds);
        }

        public ActionResult UpdateConfirmStudent(int company_id, string company_name, string student_id, string confirm_status, TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            string _update = "update Student set confirm_status = '"+ confirm_status +"' where student_id = '"+ student_id +"' ;";
            _ = db.insert_update_delete(_update);

            return RedirectToAction("ConfirmStudent", new { company_id, company_name, user_id_title = ti.user_id_title, name_title = ti.name_title});
        }




        // View Status all ------------------------------------------------------------
        public ActionResult ListStudent(TitileAll ti)
        {
            //ti.user_id_title = "kittisak";
            //ti.name_title = "นายกิตติศักดิ์ โคมเดือน";
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "List Student";
            return View();
        }



        public ActionResult ProcessStudent(TitileAll ti, string student_id)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Detail Document";

            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var _login = "SELECT* FROM Student s, Account a where student_id = '" + student_id + "' and s.student_id = a.user_id ;";
            _login += "select* from Document d, Supervisor s where d.student_id = '" + student_id + "' and d.student_id = s.student_id ;";
            _login += "select* from Work where student_id = '" + student_id + "' ;";

            ds = db.select(_login, "Student");
            if (ds.Tables["Student"].Rows[0]["name_en"].ToString() == "")
            {
                ViewData["no03"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no06"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no07"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no09"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no15"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no08"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
                ViewData["no14"] = "นักศึกษายังไม่ได้กรอกข้อมูลใช้งานเว็บไซต์";
            }
            else
            {
                if (ds.Tables["Student"].Rows[0]["name_en"].ToString() == "")
                {
                    ViewData["no03"] = "ยังไม่ได้ให้ข้อมูล";
                }
                else
                {
                    ViewData["no03"] = "เรียบร้อย";
                }

                if (ds.Tables["Student"].Rows[0]["company_id"].ToString() == "")
                {
                    ViewData["no06"] = "ยังไม่ได้เลือกสถานประการ";
                }
                else
                {
                    ViewData["no06"] = "เรียบร้อย";
                }

                if (ds.Tables["Student1"].Rows[0]["title_th"].ToString() == "" && ds.Tables["Student1"].Rows[0]["name"].ToString() == "")
                {
                    ViewData["no07"] = "ยังไม่ได้ให้ข้อมูล";
                    ViewData["no09"] = "ยังไม่ได้ให้ข้อมูล";
                    ViewData["no15"] = "ยังไม่ได้ให้ข้อมูล";
                }
                else
                {
                    ViewData["no07"] = "เรียบร้อย";
                    ViewData["no09"] = "เรียบร้อย";
                    ViewData["no15"] = "เรียบร้อย";
                }

                if (ds.Tables["Student2"].Rows.Count == 0)
                {
                    ViewData["no08"] = "ยังไม่ได้ให้ข้อมูล";
                }
                else
                {
                    ViewData["no08"] = "เรียบร้อย";
                }
            }
            ViewData["no14"] = "ไม่มีเอกสาร";
            ViewData["name_std"] = ds.Tables["Student"].Rows[0]["name_th"];
            ViewData["id"] = student_id.Substring(0,12)+"-"+student_id.Substring(12,1);
            ViewData["major"] = ds.Tables["Student"].Rows[0]["major"];
             
            return View();
        }




        //-------------------------------------------------------------------------------------------------

        public JsonResult Area(string province)
        {
            List<string> AreaList = new List<string>();
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var _Address = "SELECT DISTINCT area FROM PrototypeAddress where province = '" + province + "' order by area ;";
            ds = db.select(_Address, "PrototypeAddress");

            int _count = ds.Tables["PrototypeAddress"].Rows.Count;

            for (int i = 0; i < _count; i++)
            {
                AreaList.Add(ds.Tables["PrototypeAddress"].Rows[i]["area"].ToString());
            }

            return Json(AreaList);
        }

        public JsonResult Sub_Area(string area, string province)
        {
            List<string> Sub_AreaList = new List<string>();
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var _Address = "SELECT sub_area FROM PrototypeAddress where area = '" + area + "' and province = '" + province + "' order by sub_area ;";
            ds = db.select(_Address, "PrototypeAddress");

            int _count = ds.Tables["PrototypeAddress"].Rows.Count;

            for (int i = 0; i < _count; i++)
            {
                Sub_AreaList.Add(ds.Tables["PrototypeAddress"].Rows[i]["sub_area"].ToString());
            }

            return Json(Sub_AreaList);
        }

        public JsonResult Prototype_id(string sub_area, string area, string province)
        {
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var _Address = "SELECT prototype_id FROM PrototypeAddress where sub_area = '" + sub_area + "' and area = '" + area + "' and province = '" + province + "';";
            ds = db.select(_Address, "PrototypeAddress");

            int prototype_id = int.Parse(ds.Tables["PrototypeAddress"].Rows[0]["prototype_id"].ToString());

            return Json(prototype_id);
        }

        public JsonResult ViewStudent(string major, string years)
        {
            List<ListStudent> StudentList = new List<ListStudent>();
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var _Address = "select * from Student s, Account a where s.student_id = a.user_id and a.major = '"+ major + "' and semester_out like '%"+years+"%' ;";
            ds = db.select(_Address, "Student");

            int _count = ds.Tables["Student"].Rows.Count;

            for (int i = 0; i < _count; i++)
            {
                StudentList.Add(new ListStudent { term = ds.Tables["Student"].Rows[i]["semester_out"].ToString().Substring(5,1), student_id = ds.Tables["Student"].Rows[i]["student_id"].ToString(), name = ds.Tables["Student"].Rows[i]["name_th"].ToString() });
            }
            
            return Json(StudentList);
        }
        public int[] GetIntArray(int num)
        {
            List<int> listOfInts = new List<int>();
            while (num > 0)
            {
                listOfInts.Add(num % 10);
                num = num / 10;
            }
            listOfInts.Reverse();
            return listOfInts.ToArray();
        }

        public string thaiNumber(int[] a)
        {
            string[] thai_num = { "๐", "๑", "๒", "๓", "๔", "๕", "๖", "๗", "๘", "๙" };
            string num_thai = "";

            for (int i = 0; i < a.Length; i++)
            {
                switch (a[i].ToString())
                {
                    case "0": num_thai += thai_num[0].ToString(); break;
                    case "1": num_thai += thai_num[1].ToString(); break;
                    case "2": num_thai += thai_num[2].ToString(); break;
                    case "3": num_thai += thai_num[3].ToString(); break;
                    case "4": num_thai += thai_num[4].ToString(); break;
                    case "5": num_thai += thai_num[5].ToString(); break;
                    case "6": num_thai += thai_num[6].ToString(); break;
                    case "7": num_thai += thai_num[7].ToString(); break;
                    case "8": num_thai += thai_num[8].ToString(); break;
                    case "9": num_thai += thai_num[9].ToString(); break;

                }
            }

            return num_thai;
        }

        [HttpPost]
        public ActionResult send(string id, string getSem)
        {
            if (getSem != null)
            {
                string[] sem_split = getSem.Split('-');
                string term = sem_split[1];
                string year = sem_split[0];

                if (term == "3")
                {
                    ConnectDB db = new ConnectDB();
                    DataSet ds = new DataSet();

                    var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว';";
                    sql += "select c.company_id, c.name_th , s.student_id , s.name_th , a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว' order by c.company_id;";
                    sql += "select semester_id, date_start, date_end from DefineSemester where semester_id = '" + getSem + "';";
                    sql += "select distinct a.major ,c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว' order by c.company_id;";

                    ds = db.select(sql, "Student");

                    if (ds.Tables[3].Rows.Count == 0 || ds == null)  //// ตรวจว่ามีนักศึกษาที่ 'รับแล้ว' หรือไม่?
                    {
                        return null;
                    }
                    else
                    {
                        string fileName = "sendSummer.pdf";
                        string fileName1 = "20180924-5.pdf";
                        string sourcePath = Server.MapPath("~/pdf_file");
                        string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

                        // Use Path class to manipulate file and directory paths.
                        string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        string destFile = System.IO.Path.Combine(targetPath, id + "-" + fileName);
                        System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                        string oldFile = Server.MapPath("~/pdf_file/" + fileName);
                        string oldFile1 = Server.MapPath("~/pdf_file/" + fileName1);
                        string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + id + "-" + fileName);

                        DateTime createDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));  ///// new DateTime(2019, 6, 18)
                        DateTime startDate = Convert.ToDateTime(ds.Tables[2].Rows[0][1].ToString());
                        DateTime endDate = Convert.ToDateTime(ds.Tables[2].Rows[0][2].ToString());
                        DateTime deadLineDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));

                        string[] thai_month = { "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม" };

                        int[] ls_createDate = GetIntArray(Convert.ToInt32(createDate.Day));
                        int[] ls_createYear = GetIntArray((Convert.ToInt32(createDate.Year)) + 543);
                        int[] ls_startDate = GetIntArray(Convert.ToInt32(startDate.Day));
                        int[] ls_startYear = GetIntArray((Convert.ToInt32(startDate.Year)) + 543);
                        int[] ls_endDate = GetIntArray(Convert.ToInt32(endDate.Day));
                        int[] ls_endYear = GetIntArray((Convert.ToInt32(endDate.Year)) + 543);
                        int[] ls_deadLineDate = GetIntArray(Convert.ToInt32(deadLineDate.Day));
                        int[] ls_deadLineYear = GetIntArray((Convert.ToInt32(deadLineDate.Year)) + 543);

                        string createDate_thai = thaiNumber(ls_createDate) + " " + thai_month[Convert.ToInt32(createDate.Month) - 1] + " " + thaiNumber(ls_createYear);
                        string startDate_thai = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear);
                        string endDate_thai = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear);
                        string deadLineDate_thai = thaiNumber(ls_deadLineDate) + " " + thai_month[Convert.ToInt32(deadLineDate.Month) - 1] + " " + thaiNumber(ls_deadLineYear);

                        // open the reader
                        PdfReader reader = new PdfReader(oldFile);
                        Rectangle size = reader.GetPageSizeWithRotation(1);
                        Document document = new Document(size);

                        // open the writer
                        FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                        PdfWriter writer = PdfWriter.GetInstance(document, fs);
                        document.Open();

                        // the pdf content


                        string text = "Some random blablablabla...";
                        int total_student = 0;
                        int[] ls_totalstud;

                        for (int i = 0; i < ds.Tables[3].Rows.Count; i++)
                        {
                            document.NewPage();
                            PdfContentByte cb = writer.DirectContent;

                            // select the font properties
                            BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            cb.SetColorFill(BaseColor.DARK_GRAY);
                            cb.SetFontAndSize(bf, 15);

                            PdfImportedPage page = writer.GetImportedPage(reader, 1);
                            cb.AddTemplate(page, 0, 0);

                            cb.BeginText();

                            text = "ผู้จัดการฝ่ายบุคคล " + ds.Tables[3].Rows[i][2].ToString();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 124, 573, 0);

                            text = createDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 316, 627, 0);

                            text = "ภาควิชา" + ds.Tables[3].Rows[i][0].ToString() + " สังกัดคณะวิศวกรรมศาสตร์";
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 273, 513, 0);

                            ls_totalstud = GetIntArray(Convert.ToInt32(year));
                            text = thaiNumber(ls_totalstud);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 509, 492, 0);

                            ls_totalstud = GetIntArray(Convert.ToInt32(startDate.Day));
                            text = thaiNumber(ls_totalstud) + thai_month[Convert.ToInt32(createDate.Month) - 1];
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 461, 438, 0);

                            ls_totalstud = GetIntArray(Convert.ToInt32(startDate.Year));
                            text = thaiNumber(ls_totalstud) + " ถึง วันที่ " + deadLineDate_thai + " โดยมีรายชื่อดังนี้";
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 85, 417, 0);

                            cb.EndText();

                            ////////////////////////////////////////////////////////////////////  นับจำนวน นศ.

                            for (int j = 0; j < ds.Tables[1].Rows.Count; j++)
                            {
                                if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[j][0].ToString() && ds.Tables[1].Rows[j][4].ToString() == ds.Tables[3].Rows[i][0].ToString())  //// if (ds.Tables[0].Rows[i][0].ToString() == ds.Tables[1].Rows[j][0].ToString())
                                {
                                    total_student++;
                                }
                            }

                            /////////////////////////////////////////////////////////////////////

                            if (total_student != 0 && total_student <= 5)  /// ถ้า นศ. ไม่เกิน 5 คน
                            {
                                PdfReader reader1 = new PdfReader(oldFile1);
                                PdfImportedPage importedPage1 = writer.GetImportedPage(reader1, 1);

                                document.NewPage();

                                BaseFont baseFont1 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                                PdfContentByte cb1 = writer.DirectContent;
                                cb1.SetColorFill(BaseColor.DARK_GRAY);
                                cb1.SetFontAndSize(baseFont1, 12);
                                cb1.AddTemplate(importedPage1, 0, 0);

                                int buntud = 530;  ////// บรรทัด
                                int run_number = 0;

                                cb1.BeginText();
                                text = ds.Tables[3].Rows[i][2].ToString();    ///// text = ds.Tables[0].Rows[i][1].ToString();
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 145, 661, 0);
                                cb1.EndText();

                                for (int j = 0; j < ds.Tables[1].Rows.Count; j++)
                                {
                                    if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[j][0].ToString() && ds.Tables[1].Rows[j][4].ToString() == ds.Tables[3].Rows[i][0].ToString())  //// if (ds.Tables[0].Rows[i][0].ToString() == ds.Tables[1].Rows[j][0].ToString())
                                    {
                                        run_number++;
                                        ls_totalstud = GetIntArray(run_number);

                                        cb1.BeginText();

                                        text = thaiNumber(ls_totalstud);
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 72, buntud, 0);

                                        text = ds.Tables[1].Rows[j][3].ToString();
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 114, buntud, 0);

                                        text = ds.Tables[1].Rows[j][4].ToString();
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 335, buntud, 0);

                                        text = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear) + " ถึง";
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 401, buntud + 5, 0);

                                        text = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear);
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 401, buntud - 5, 0);

                                        cb1.EndText();

                                        buntud = buntud - 27;
                                    }
                                }

                                total_student = 0;
                            }
                            else //// ถ้า นศ.เกิน 5 คน
                            {
                                PdfReader reader1 = new PdfReader(oldFile1);
                                PdfImportedPage importedPage1 = writer.GetImportedPage(reader1, 1);

                                document.NewPage();

                                BaseFont baseFont1 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                                PdfContentByte cb1 = writer.DirectContent;
                                cb1.SetColorFill(BaseColor.DARK_GRAY);
                                cb1.SetFontAndSize(baseFont1, 12);
                                cb1.AddTemplate(importedPage1, 0, 0);

                                cb1.BeginText();
                                text = ds.Tables[3].Rows[i][2].ToString();    ///// text = ds.Tables[0].Rows[i][1].ToString();
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 145, 661, 0);
                                cb1.EndText();

                                cb1.BeginText();

                                text = "รายชื่อตามแนบ";
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 175, 530, 0);
                                cb1.EndText();


                                document.NewPage(); //// หน้าตารางรายชื่อตามแนบ

                                BaseFont bfTimes2 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                Font sarabun = new Font(bfTimes2, 15);
                                sarabun.SetColor(0, 0, 0);

                                ////////////////////////////////////////////////////////////////////////////////

                                PdfPTable table = new PdfPTable(5);
                                float[] widths = new float[] { 18f, 60f, 65f, 55f, 25f };
                                int run_number = 0;

                                table.TotalWidth = 500f;
                                table.LockedWidth = true;
                                table.SetWidths(widths);

                                PdfPCell cell = new PdfPCell(new Phrase("ลำดับที่", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell2 = new PdfPCell(new Phrase("ชื่อ-นามสกุลนักศึกษา", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell3 = new PdfPCell(new Phrase("สาขาวิชา", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell4 = new PdfPCell(new Phrase("ปฏิบัติงานระหว่างวันที่", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell5 = new PdfPCell(new Phrase("หมายเหตุ", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };

                                table.AddCell(cell);
                                table.AddCell(cell2);
                                table.AddCell(cell3);
                                table.AddCell(cell4);
                                table.AddCell(cell5);

                                cell.BorderColorBottom = new BaseColor(0, 0, 0);
                                cell.BorderColorTop = new BaseColor(0, 0, 0);
                                cell.BorderColorRight = new BaseColor(0, 0, 0);
                                cell.BorderColorLeft = new BaseColor(0, 0, 0);

                                for (int m = 0; m < ds.Tables[1].Rows.Count; m++)
                                {
                                    if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[m][0].ToString() && ds.Tables[1].Rows[m][4].ToString() == ds.Tables[3].Rows[i][0].ToString())
                                    {
                                        run_number++; /// เลขลำดับ
                                        ls_totalstud = GetIntArray(run_number);

                                        cell = new PdfPCell(new Phrase(thaiNumber(ls_totalstud), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                        cell2 = new PdfPCell(new Phrase(ds.Tables[1].Rows[m][3].ToString(), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT };
                                        cell3 = new PdfPCell(new Phrase(ds.Tables[1].Rows[m][4].ToString(), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                        cell4 = new PdfPCell(new Phrase(thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear) + " ถึง " + thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT }; ;
                                        cell5 = new PdfPCell(new Phrase("", sarabun));

                                        cell.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell2.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell3.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell4.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell5.BorderColorBottom = new BaseColor(255, 255, 255);

                                        table.AddCell(cell);
                                        table.AddCell(cell2);
                                        table.AddCell(cell3);
                                        table.AddCell(cell4);
                                        table.AddCell(cell5);

                                    }

                                }

                                //table.WriteSelectedRows(50, 100, 200, 50, cb1);

                                document.Add(table);

                            }
                        }


                        document.Close();
                        fs.Close();
                        writer.Close();
                        reader.Close();
                        return RedirectToAction("GetReport", "Authorities", new { doc_id = id + "-" + fileName });
                        //return null;
                    }
                }
                else
                {
                    ConnectDB db = new ConnectDB();
                    DataSet ds = new DataSet();

                    var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว';";
                    sql += "select c.company_id, c.name_th , s.student_id , s.name_th , a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว' order by c.company_id;";
                    sql += "select semester_id, date_start, date_end from DefineSemester where semester_id = '" + getSem + "';";
                    sql += "select distinct a.major ,c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รับแล้ว' order by c.company_id;";

                    ds = db.select(sql, "Student");

                    if (ds.Tables[3].Rows.Count == 0 || ds == null)  //// ตรวจว่ามีนักศึกษาที่ 'รับแล้ว' หรือไม่?
                    {
                        return null;
                    }
                    else
                    {
                        string fileName = "send.pdf";
                        string fileName1 = "20180924-5.pdf";
                        string sourcePath = Server.MapPath("~/pdf_file");
                        string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

                        // Use Path class to manipulate file and directory paths.
                        string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        string destFile = System.IO.Path.Combine(targetPath, id + "-" + fileName);
                        System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                        string oldFile = Server.MapPath("~/pdf_file/" + fileName);
                        string oldFile1 = Server.MapPath("~/pdf_file/" + fileName1);
                        string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + id + "-" + fileName);

                        DateTime createDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));  ///// new DateTime(2019, 6, 18)
                        DateTime startDate = Convert.ToDateTime(ds.Tables[2].Rows[0][1].ToString());
                        DateTime endDate = Convert.ToDateTime(ds.Tables[2].Rows[0][2].ToString());
                        DateTime deadLineDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));

                        string[] thai_month = { "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม" };

                        int[] ls_createDate = GetIntArray(Convert.ToInt32(createDate.Day));
                        int[] ls_createYear = GetIntArray((Convert.ToInt32(createDate.Year)) + 543);
                        int[] ls_startDate = GetIntArray(Convert.ToInt32(startDate.Day));
                        int[] ls_startYear = GetIntArray((Convert.ToInt32(startDate.Year)) + 543);
                        int[] ls_endDate = GetIntArray(Convert.ToInt32(endDate.Day));
                        int[] ls_endYear = GetIntArray((Convert.ToInt32(endDate.Year)) + 543);
                        int[] ls_deadLineDate = GetIntArray(Convert.ToInt32(deadLineDate.Day));
                        int[] ls_deadLineYear = GetIntArray((Convert.ToInt32(deadLineDate.Year)) + 543);



                        string createDate_thai = thaiNumber(ls_createDate) + " " + thai_month[Convert.ToInt32(createDate.Month) - 1] + " " + thaiNumber(ls_createYear);
                        string startDate_thai = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1];
                        string endDate_thai = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear);
                        string deadLineDate_thai = thaiNumber(ls_deadLineDate) + " " + thai_month[Convert.ToInt32(deadLineDate.Month) - 1] + " " + thaiNumber(ls_deadLineYear);


                        // open the reader
                        PdfReader reader = new PdfReader(oldFile);
                        Rectangle size = reader.GetPageSizeWithRotation(1);
                        Document document = new Document(size);

                        // open the writer
                        FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                        PdfWriter writer = PdfWriter.GetInstance(document, fs);
                        document.Open();

                        string text = "Some random blablablabla...";
                        int total_student = 0;
                        int[] ls_totalstud;

                        for (int i = 0; i < ds.Tables[3].Rows.Count; i++) //// for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            document.NewPage();

                            BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            PdfImportedPage importedPage = writer.GetImportedPage(reader, 1);

                            PdfContentByte cb = writer.DirectContent;
                            cb.SetColorFill(BaseColor.DARK_GRAY);
                            cb.SetFontAndSize(baseFont, 15);
                            cb.AddTemplate(importedPage, 0, 0);

                            cb.BeginText();

                            text = "ผู้จัดการฝ่ายบุคคล " + ds.Tables[3].Rows[i][2].ToString();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 124, 573, 0);

                            text = createDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 316, 627, 0);

                            text = "ภาควิชา" + ds.Tables[3].Rows[i][0].ToString() + " สังกัดคณะวิศวกรรมศาสตร์";  /// "ภาควิชา"
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 273, 513, 0);

                            ls_totalstud = GetIntArray(Convert.ToInt32(term));
                            text = thaiNumber(ls_totalstud);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 416, 492, 0);

                            ls_totalstud = GetIntArray(Convert.ToInt32(year));
                            text = thaiNumber(ls_totalstud); ;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 85, 471, 0);

                            ////////////////////////////////////////////////////////////////////  นับจำนวน นศ.

                            for (int j = 0; j < ds.Tables[1].Rows.Count; j++)
                            {
                                if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[j][0].ToString() && ds.Tables[1].Rows[j][4].ToString() == ds.Tables[3].Rows[i][0].ToString())  //// if (ds.Tables[0].Rows[i][0].ToString() == ds.Tables[1].Rows[j][0].ToString())
                                {
                                    total_student++;
                                }
                            }

                            /////////////////////////////////////////////////////////////////////

                            ls_totalstud = GetIntArray(total_student);

                            text = thaiNumber(ls_totalstud);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 281, 471, 0);

                            text = startDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 461, 438, 0);

                            text = thaiNumber(ls_startYear) + " ถึง " + endDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 85, 417, 0);

                            cb.EndText();


                            if (total_student != 0 && total_student <= 5)  /// ถ้า นศ. ไม่เกิน 5 คน
                            {
                                PdfReader reader1 = new PdfReader(oldFile1);
                                PdfImportedPage importedPage1 = writer.GetImportedPage(reader1, 1);

                                document.NewPage();

                                BaseFont baseFont1 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                                PdfContentByte cb1 = writer.DirectContent;
                                cb1.SetColorFill(BaseColor.DARK_GRAY);
                                cb1.SetFontAndSize(baseFont, 12);
                                cb1.AddTemplate(importedPage1, 0, 0);

                                int buntud = 530;  ////// บรรทัด
                                int run_number = 0;

                                cb1.BeginText();
                                text = ds.Tables[3].Rows[i][2].ToString();    ///// text = ds.Tables[0].Rows[i][1].ToString();
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 145, 661, 0);
                                cb1.EndText();

                                for (int j = 0; j < ds.Tables[1].Rows.Count; j++)
                                {
                                    if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[j][0].ToString() && ds.Tables[1].Rows[j][4].ToString() == ds.Tables[3].Rows[i][0].ToString())  //// if (ds.Tables[0].Rows[i][0].ToString() == ds.Tables[1].Rows[j][0].ToString())
                                    {
                                        run_number++;
                                        ls_totalstud = GetIntArray(run_number);

                                        cb1.BeginText();

                                        text = thaiNumber(ls_totalstud);
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 72, buntud, 0);

                                        text = ds.Tables[1].Rows[j][3].ToString();
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 114, buntud, 0);

                                        text = ds.Tables[1].Rows[j][4].ToString();
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 335, buntud, 0);

                                        text = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear) + " ถึง";
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 401, buntud + 5, 0);

                                        text = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear);
                                        cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 401, buntud - 5, 0);

                                        cb1.EndText();

                                        buntud = buntud - 27;
                                    }
                                }

                                total_student = 0;
                            }
                            else //// ถ้า นศ.เกิน 5 คน
                            {
                                PdfReader reader1 = new PdfReader(oldFile1);
                                PdfImportedPage importedPage1 = writer.GetImportedPage(reader1, 1);

                                document.NewPage();

                                BaseFont baseFont1 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                                PdfContentByte cb1 = writer.DirectContent;
                                cb1.SetColorFill(BaseColor.DARK_GRAY);
                                cb1.SetFontAndSize(baseFont, 12);
                                cb1.AddTemplate(importedPage1, 0, 0);

                                cb1.BeginText();
                                text = ds.Tables[3].Rows[i][2].ToString();    ///// text = ds.Tables[0].Rows[i][1].ToString();
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 145, 661, 0);
                                cb1.EndText();

                                cb1.BeginText();

                                text = "รายชื่อตามแนบ";
                                cb1.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 175, 530, 0);
                                cb1.EndText();


                                document.NewPage(); //// หน้าตารางรายชื่อตามแนบ

                                BaseFont bfTimes2 = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                Font sarabun = new Font(bfTimes2, 15);
                                sarabun.SetColor(0, 0, 0);

                                ////////////////////////////////////////////////////////////////////////////////

                                PdfPTable table = new PdfPTable(5);
                                float[] widths = new float[] { 18f, 60f, 65f, 55f, 25f };
                                int run_number = 0;

                                table.TotalWidth = 500f;
                                table.LockedWidth = true;
                                table.SetWidths(widths);

                                PdfPCell cell = new PdfPCell(new Phrase("ลำดับที่", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell2 = new PdfPCell(new Phrase("ชื่อ-นามสกุลนักศึกษา", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell3 = new PdfPCell(new Phrase("สาขาวิชา", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell4 = new PdfPCell(new Phrase("ปฏิบัติงานระหว่างวันที่", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                PdfPCell cell5 = new PdfPCell(new Phrase("หมายเหตุ", sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };

                                table.AddCell(cell);
                                table.AddCell(cell2);
                                table.AddCell(cell3);
                                table.AddCell(cell4);
                                table.AddCell(cell5);

                                cell.BorderColorBottom = new BaseColor(0, 0, 0);
                                cell.BorderColorTop = new BaseColor(0, 0, 0);
                                cell.BorderColorRight = new BaseColor(0, 0, 0);
                                cell.BorderColorLeft = new BaseColor(0, 0, 0);

                                for (int m = 0; m < ds.Tables[1].Rows.Count; m++)
                                {
                                    if (ds.Tables[3].Rows[i][1].ToString() == ds.Tables[1].Rows[m][0].ToString() && ds.Tables[1].Rows[m][4].ToString() == ds.Tables[3].Rows[i][0].ToString())
                                    {
                                        run_number++; /// เลขลำดับ
                                        ls_totalstud = GetIntArray(run_number);

                                        cell = new PdfPCell(new Phrase(thaiNumber(ls_totalstud), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                        cell2 = new PdfPCell(new Phrase(ds.Tables[1].Rows[m][3].ToString(), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT };
                                        cell3 = new PdfPCell(new Phrase(ds.Tables[1].Rows[m][4].ToString(), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER };
                                        cell4 = new PdfPCell(new Phrase(thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear) + " ถึง " + thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear), sarabun)) { VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT }; ;
                                        cell5 = new PdfPCell(new Phrase("", sarabun));

                                        cell.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell2.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell3.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell4.BorderColorBottom = new BaseColor(255, 255, 255);
                                        cell5.BorderColorBottom = new BaseColor(255, 255, 255);

                                        table.AddCell(cell);
                                        table.AddCell(cell2);
                                        table.AddCell(cell3);
                                        table.AddCell(cell4);
                                        table.AddCell(cell5);

                                    }

                                }

                                //table.WriteSelectedRows(50, 100, 200, 50, cb1);

                                document.Add(table);

                            }

                        }

                        document.Close();
                        fs.Close();
                        writer.Close();
                        reader.Close();
                        return RedirectToAction("GetReport", "Authorities", new { doc_id = id + "-" + fileName });
                    }
                    //return null;
                }
            }
            else
            {
                return null;
            }
        }

        [HttpPost]
        public ActionResult re(string id, string getSem)
        {
            if (getSem != null)
            {
                string[] sem_split = getSem.Split('-');
                string term = sem_split[1];
                string year = sem_split[0];

                if (term == "3")
                {
                    ConnectDB db = new ConnectDB();
                    DataSet ds = new DataSet();

                    var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รอการตอบรับ';";
                    sql += "select c.company_id, c.name_th , s.student_id , s.name_th , a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รอการตอบรับ' order by c.company_id;";
                    sql += "select semester_id, date_start, date_end from DefineSemester where semester_id = '" + getSem + "';";

                    ds = db.select(sql, "Student");

                    if (ds.Tables[0].Rows.Count == 0 || ds == null) ////// ตรวจว่ามีเด็กที่ 'รอการตอบรับ' หรือไม่?
                    {
                        return null;
                    }
                    else
                    {
                        string fileName = "reSummer.pdf";
                        string sourcePath = Server.MapPath("~/pdf_file");
                        string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

                        // Use Path class to manipulate file and directory paths.
                        string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        string destFile = System.IO.Path.Combine(targetPath, id + "-" + fileName);
                        System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                        string oldFile = Server.MapPath("~/pdf_file/" + fileName);
                        string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + id + "-" + fileName);

                        DateTime createDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));  ///// new DateTime(2019, 6, 18)
                        DateTime startDate = Convert.ToDateTime(ds.Tables[2].Rows[0][1].ToString());
                        DateTime endDate = Convert.ToDateTime(ds.Tables[2].Rows[0][2].ToString());
                        DateTime deadLineDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));

                        string[] thai_month = { "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม" };

                        int[] ls_createDate = GetIntArray(Convert.ToInt32(createDate.Day));
                        int[] ls_createYear = GetIntArray((Convert.ToInt32(createDate.Year)) + 543);
                        int[] ls_startDate = GetIntArray(Convert.ToInt32(startDate.Day));
                        int[] ls_startYear = GetIntArray((Convert.ToInt32(startDate.Year)) + 543);
                        int[] ls_endDate = GetIntArray(Convert.ToInt32(endDate.Day));
                        int[] ls_endYear = GetIntArray((Convert.ToInt32(endDate.Year)) + 543);
                        int[] ls_deadLineDate = GetIntArray(Convert.ToInt32(deadLineDate.Day));
                        int[] ls_deadLineYear = GetIntArray((Convert.ToInt32(deadLineDate.Year)) + 543);



                        string createDate_thai = thaiNumber(ls_createDate) + " " + thai_month[Convert.ToInt32(createDate.Month) - 1] + " " + thaiNumber(ls_createYear);
                        string startDate_thai = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear);
                        string endDate_thai = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1];
                        string deadLineDate_thai = thaiNumber(ls_deadLineDate) + " " + thai_month[Convert.ToInt32(deadLineDate.Month) - 1] + " " + thaiNumber(ls_deadLineYear);


                        // open the reader
                        PdfReader reader = new PdfReader(oldFile);
                        Rectangle size = reader.GetPageSizeWithRotation(1);
                        Document document = new Document(size);

                        // open the writer
                        FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                        PdfWriter writer = PdfWriter.GetInstance(document, fs);
                        document.Open();

                        // the pdf content
                        PdfContentByte cb = writer.DirectContent;

                        // select the font properties

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            document.NewPage();
                            BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            cb.SetColorFill(BaseColor.DARK_GRAY);
                            cb.SetFontAndSize(bf, 15);

                            PdfImportedPage page = writer.GetImportedPage(reader, 1);
                            cb.AddTemplate(page, 0, 0);

                            string text = "Some random blablablabla...";

                            cb.BeginText();

                            text = "ผู้จัดการฝ่ายบุคคล " + ds.Tables[0].Rows[i][1].ToString();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 110, 520, 0);

                            //text = "รายชื่อตามแนบ";
                            //cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 309, 313, 0);

                            text = createDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 309, 579, 0);

                            text = startDate_thai + " ถึง " + endDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 378, 361, 0);

                            text = thaiNumber(ls_endYear);
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 72, 340, 0);

                            text = deadLineDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 252, 265, 0);

                            cb.EndText();
                        }

                        document.Close();
                        fs.Close();
                        writer.Close();
                        reader.Close();

                        return RedirectToAction("GetReport", "Authorities", new { doc_id = id + "-" + fileName });
                    }
                }
                else
                {
                    ConnectDB db = new ConnectDB();
                    DataSet ds = new DataSet();

                    var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รอการตอบรับ';";
                    sql += "select c.company_id, c.name_th , s.student_id , s.name_th , a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where s.semester_out = '" + getSem + "' and s.confirm_status = 'รอการตอบรับ' order by c.company_id;";
                    sql += "select semester_id, date_start, date_end from DefineSemester where semester_id = '" + getSem + "';";

                    ds = db.select(sql, "Student");

                    if (ds.Tables[0].Rows.Count == 0 || ds == null) ////// ตรวจว่ามีเด็กที่ 'รอการตอบรับ' หรือไม่?
                    {
                        return null;
                    }
                    else
                    {
                        string fileName = "re.pdf";
                        string sourcePath = Server.MapPath("~/pdf_file");
                        string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

                        // Use Path class to manipulate file and directory paths.
                        string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                        string destFile = System.IO.Path.Combine(targetPath, id + "-" + fileName);
                        System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                        string oldFile = Server.MapPath("~/pdf_file/" + fileName);
                        string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + id + "-" + fileName);

                        DateTime createDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                        DateTime startDate = Convert.ToDateTime(ds.Tables[2].Rows[0][1].ToString());
                        DateTime endDate = Convert.ToDateTime(ds.Tables[2].Rows[0][2].ToString());
                        DateTime deadLineDate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));

                        string[] thai_month = { "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน", "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม" };

                        int[] ls_createDate = GetIntArray(Convert.ToInt32(createDate.Day));
                        int[] ls_createYear = GetIntArray((Convert.ToInt32(createDate.Year)) + 543);
                        int[] ls_startDate = GetIntArray(Convert.ToInt32(startDate.Day));
                        int[] ls_startYear = GetIntArray((Convert.ToInt32(startDate.Year)) + 543);
                        int[] ls_endDate = GetIntArray(Convert.ToInt32(endDate.Day));
                        int[] ls_endYear = GetIntArray((Convert.ToInt32(endDate.Year)) + 543);
                        int[] ls_deadLineDate = GetIntArray(Convert.ToInt32(deadLineDate.Day));
                        int[] ls_deadLineYear = GetIntArray((Convert.ToInt32(deadLineDate.Year)) + 543);



                        string createDate_thai = thaiNumber(ls_createDate) + " " + thai_month[Convert.ToInt32(createDate.Month) - 1] + " " + thaiNumber(ls_createYear);
                        string startDate_thai = thaiNumber(ls_startDate) + " " + thai_month[Convert.ToInt32(startDate.Month) - 1] + " " + thaiNumber(ls_startYear);
                        string endDate_thai = thaiNumber(ls_endDate) + " " + thai_month[Convert.ToInt32(endDate.Month) - 1] + " " + thaiNumber(ls_endYear);
                        string deadLineDate_thai = thaiNumber(ls_deadLineDate) + " " + thai_month[Convert.ToInt32(deadLineDate.Month) - 1] + " " + thaiNumber(ls_deadLineYear);


                        // open the reader
                        PdfReader reader = new PdfReader(oldFile);
                        Rectangle size = reader.GetPageSizeWithRotation(1);
                        Document document = new Document(size);

                        // open the writer
                        FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                        PdfWriter writer = PdfWriter.GetInstance(document, fs);
                        document.Open();

                        // the pdf content
                        PdfContentByte cb = writer.DirectContent;


                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            string text = "Some random blablablabla...";

                            document.NewPage();
                            //PdfContentByte cb = writer.DirectContent;
                            BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            cb.SetColorFill(BaseColor.DARK_GRAY);
                            cb.SetFontAndSize(bf, 15);
                            PdfImportedPage page = writer.GetImportedPage(reader, 1);
                            cb.AddTemplate(page, 0, 0);

                            cb.BeginText();

                            text = "ผู้จัดการฝ่ายบุคคล " + ds.Tables[0].Rows[i][1].ToString();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 107, 558, 0);

                            text = createDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 311, 615, 0);


                            text = "วันที่ " + startDate_thai + " ถึงวันที่ " + endDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 72, 330, 0);

                            text = "ภายในวันที่ " + deadLineDate_thai;
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 328, 291, 0);

                            cb.EndText();
                        }

                        document.Close();
                        fs.Close();
                        writer.Close();
                        reader.Close();

                        return RedirectToAction("GetReport", "Authorities", new { doc_id = id + "-" + fileName });
                    }
                }
            }
            else
            {
                return null;
            }

        }

        public ActionResult sendBook(TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var sql = "select semester_id from DefineSemester order by semester_id desc;";

            ds = db.select(sql, "DefineSemester");

            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;

            return View(ds);
            //return RedirectToAction("send", "Authorities", new { company_id = id });
        }

        public ActionResult reBook(TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();

            var sql = "select semester_id from DefineSemester order by semester_id desc;";

            ds = db.select(sql, "DefineSemester");

            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;

            return View(ds);
            //return RedirectToAction("send", "Authorities", new { company_id = id });
        }

        public FileResult GetReport(string doc_id) // เปิด PDF ในหน้าใหม่
        {

            string ReportURL = Server.MapPath("~/pdf_file/pdf_copy/" + doc_id);
            byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
            string physicalPath = Server.MapPath("~/pdf_file/pdf_copy/" + doc_id);

            var uri = new Uri(physicalPath, UriKind.Absolute); //ลบไฟล์
            System.IO.File.Delete(uri.LocalPath);
            return File(FileBytes, "application/pdf");
        }

        public ActionResult Report(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;

            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            //DefineSemester def_s = new DefineSemester();
            string sql = "SELECT semester_id FROM DefineSemester order by semester_id desc;";
            sql += "select distinct a.major from Student s inner join Account a on a.user_id = s.student_id ;";
            ds = db.select(sql, "Semester");

            return View(ds);
        }

        public ActionResult excelReport(string sem, string major)
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

    public class ListStudent
    {
        public string term { get; set; }
        public string student_id { get; set; }
        public string name { get; set; }
    }

}


