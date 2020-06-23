using iTextSharp.text;
using iTextSharp.text.pdf;
using Project_Test1.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Test1.Controllers
{
    public class StudentController : Controller
    {
        // GET: Student
        public static string st_map_pic;

        public ActionResult logout()
        {
            st_map_pic = null;
            return RedirectToAction("LoginTemplate", "Home");
        }

        public ActionResult Index(TitileAll ti)
        {
            //ti.user_id_title = "1160304620454";
            //ViewBag.Message = status;

            ConnectDB db = new ConnectDB();

            var _login = "SELECT * FROM Student where student_id = '" + ti.user_id_title + "' AND name_en IS NOT NULL ;";
            _login += "select * from Company where company_id in (SELECT company_id FROM Student where student_id = '" + ti.user_id_title + "') ;";
            _login += "select * from `match` m, TeachersandAuthorities t where student_id = '" + ti.user_id_title + "' and m.teacher_id = t.teacher_id ;";
            _login += "select * from Document where student_id = '" + ti.user_id_title + "' ;";
            _login += "select * from Supervisor where student_id = '" + ti.user_id_title + "' ;";
            _login += "select * from Work where student_id = '" + ti.user_id_title + "' ;";
            DataSet ds = new DataSet();
            ds = db.select(_login, "Student");
            var a = ds.Tables["Student"].Rows.Count;
            if (a == 1)
            {
                if (ds.Tables["Student1"].Rows.Count == 0)
                {
                    ds.Tables["Student1"].Rows.Add();
                    ds.Tables["Student1"].Rows[0]["name_th"] = "ยังไม่ได้เลือก";
                    ds.Tables["Student"].Rows[0]["confirm_status"] = "ยังไม่ได้เลือก";
                }
                else
                {
                    if (ds.Tables["Student1"].Rows[0]["name_th"].ToString().Length < 10)
                    {
                        ds.Tables["Student1"].Rows[0]["name_th"] = ds.Tables["Student1"].Rows[0]["name_en"];
                    }
                }

                if (ds.Tables["Student2"].Rows.Count == 0)
                {
                    ds.Tables["Student2"].Rows.Add();
                    ds.Tables["Student2"].Rows[0]["name"] = "ยังไม่ระบุ";
                    ds.Tables["Student2"].Rows[0]["date_supervision"] = "ยังไม่ระบุ";
                }

                if (ds.Tables["Student3"].Rows[0]["title_th"].ToString() == "")
                {
                    ViewData["no09"] = "ยังไม่ได้ให้ข้อมูล สก. 09";
                }
                else
                {
                    ViewData["no09"] = "เรียบร้อย";
                }
                if (ds.Tables["Student4"].Rows[0]["name"].ToString() == "")
                {
                    ViewData["no07"] = "ยังไม่ได้ให้ข้อมูล สก. 07";
                }
                else
                {
                    ViewData["no07"] = "เรียบร้อย";
                }
                if (ds.Tables["Student5"].Rows.Count == 0)
                {
                    ViewData["no08"] = "ยังไม่ได้ให้ข้อมูล สก. 08";
                }
                else
                {
                    ViewData["no08"] = "เรียบร้อย";
                }

                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ds.Tables["Student"].Rows[0]["name_th"];
                ViewData["msgAlert"] = ti.msgAlert;
                ViewBag.Page = "View";

                ViewData["company_name"] = ds.Tables["Student1"].Rows[0]["name_th"];
                ViewData["status"] = ds.Tables["Student"].Rows[0]["confirm_status"];

                ViewData["name_teacher"] = ds.Tables["Student2"].Rows[0]["name"];
                ViewData["date"] = ds.Tables["Student2"].Rows[0]["date_supervision"];

                return View();
            }
            else
            {
                return RedirectToAction("FirstTime", "Student", new { student_id = ti.user_id_title });
            } 
        }



        public ActionResult EditProfile(TitileAll ti)
        {
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Edit Profile";
            ViewData["msgAlert"] = ti.msgAlert;
            ConnectDB db = new ConnectDB();
            var _edit = "select * from Student s, Account a where s.student_id = '"+ ti.user_id_title + "' and s.student_id = a.user_id ;";
            _edit += "select * from Address a, PrototypeAddress p where a.student_company_id = '" + ti.user_id_title + "' and p.prototype_id = a.prototype_id ;";
            _edit += "select * from Language where Student_id = '" + ti.user_id_title + "' ;";
            _edit += "select * from HistoryofEducation where Student_id = '" + ti.user_id_title + "' ;";
            _edit += "select * from Family where Student_id = '" + ti.user_id_title + "' ;";
            _edit += "SELECT DISTINCT province FROM PrototypeAddress where not prototype_id = 0 order by province ;";
            _edit += "select * from Family where Student_id = '" + ti.user_id_title + "' and relation = 'พี่-น้อง' ;";

            _edit += "select semester_id from DefineSemester;";
            _edit += "select confirm_status,company_id from Student where student_id = '" + ti.user_id_title + "';";
            DataSet ds = new DataSet();
            ds = db.select(_edit, "Edit");
            if (ds.Tables[8].Rows[0][0].ToString() == "รอการตอบรับ" || ds.Tables[8].Rows[0][0].ToString() == "รับแล้ว")
            {
                ViewData["check"] = "Yes";
            }

            List<string> ListProvince = new List<string>();
            int _countProvince = ds.Tables["Edit5"].Rows.Count;
            for (int i = 0; i < _countProvince; i++)
            {
                ListProvince.Add(ds.Tables["Edit5"].Rows[i]["province"].ToString());
            }
            ViewBag.province = ListProvince;

            return View(ds);
        }
        [HttpPost]
        public ActionResult EditProfile(Student std, TitileAll ti, string checkYes, string checkNo, string sem, string changeS)
        {
            ConnectDB db = new ConnectDB();
            var _edit = "select * from Student s, Account a where s.student_id = '" + ti.user_id_title + "' and s.student_id = a.user_id ;";
            _edit += "select * from HistoryofEducation where Student_id = '" + ti.user_id_title + "' ;";
            DataSet ds = new DataSet();
            ds = db.select(_edit, "Edit");

            DateTime date_birth = Convert.ToDateTime(std.date_birth);
            std.age = DateTime.Now.Year - date_birth.Year;
            if (date_birth.Month > DateTime.Now.Month) std.age = std.age - 1;
            else if (date_birth.Month == DateTime.Now.Month)
            {
                if (date_birth.Day > DateTime.Now.Day) std.age = std.age - 1;
            }

            var _update = "UPDATE Student SET name_en = '" + std.name_en + "', study = '" + std.study + "', year = " + std.year + ", advisor = '" + std.advisor + "', gpa = " + std.gpa + ", "
                + "gpax = " + std.gpax + ", identification = '" + std.identification + "', at = '" + std.at + "', date = '" + std.date + "', telephone = '" + std.telephone + "', mobile_phone = '"
                + std.mobile_phone + "', fax = '" + std.fax + "', email = '" + std.email + "', race = '" + std.race + "', nationality = '" + std.nationality + "', religion = '" + std.religion + "', "
                + "date_birth = '" + std.date_birth + "', age = " + std.age + ", gender = '" + std.gender + "', height = " + std.height + ", weight = " + std.weight + ", disease = '" + std.disease
                + "', ability_spectify = '" + std.ability_spectify + "', experience = '" + std.experience + "' WHERE student_id = '" + std.student_id + "' ;";
            for(int i=0; i<4; i++){
                _update += "update Address set prototype_id = "+std.ad_prototype_id[i]+", number = '"+std.ad_number[i]+ "', building = '" + std.ad_building[i] + "', floor = '" + std.ad_floor[i] +
                    "', lane = '" + std.ad_lane[i] + "', road = '" + std.ad_road[i] + "' where student_company_id = '" + std.student_id + "' and who = '"+std.ad_who[i]+"' ;";
            }
            for (int i=0; i<std.la_name.Length; i++)
            {
                _update += "update Language set listening = '"+std.la_listening[i]+"', reading = '"+std.la_reading[i]+"', speaking = '"+std.la_speaking[i]+"', writing = '"+std.la_writing[i]+
                    "' where student_id = '" + std.student_id + "' and name = '"+std.la_name[i]+"' ;";
            }
            for (int i = 0; i < std.hi_name.Length; i++)
            {
                _update += "update HistoryofEducation set name = '" + std.hi_name[i] + "', year = " + std.hi_year[i] + ", certificate = '" + std.hi_certificate[i] + "', major = '" + std.hi_major[i] +
                    "', gpa = " + std.hi_gpa[i] + " where student_id = '" + std.student_id + "' and number = " + int.Parse(ds.Tables["Edit1"].Rows[i]["number"].ToString()) + " ;";
            }
            _update += "update Family set name = '" + std.fa_name[0] + "', occupation = '" + std.fa_occupation[0] + "', telephone = '" + std.fa_telephone[0] + "', fax = '" + std.fa_fax[0] +
                    "', mobile_phone = '" + std.fa_mobile_phone[0] + "', email = '" + std.fa_email[0] + "', place_work = '" + std.fa_place_work[0] + "' where student_id = '" + std.student_id + "' and number = 1 ;";
            for (int i = 1; i <= 2; i++)
            {
                _update += "update Family set name = '" + std.fa_name[i] + "', occupation = '" + std.fa_occupation[i] + "', age = " + std.fa_age[i-1] + ", mobile_phone = '" + std.fa_mobile_phone[i-1] + 
                    "' where student_id = '" + std.student_id + "' and number = "+(i+1)+" ;";
            }
            for (int i = 3; i < std.fa_name.Length; i++)
            {
                _update += "update Family set name = '" + std.fa_name[i] + "', occupation = '" + std.fa_occupation[i] + "', age = " + std.fa_age[i-1] + ", position = '" + std.fa_position[i-3] +
                    "' where student_id = '" + std.student_id + "' and number = " + (i + 1) + " ;";
            }

            if(changeS == "checkYes")  /// ประสงค์ออกฝึก
            {
                if(ds.Tables[0].Rows[0]["confirm_status"].ToString() == "ไม่ประสงค์ออกฝึก")
                {
                    _update += "UPDATE Student SET confirm_status = null where student_id = '"+ std.student_id + "';";
                }
                
            }

            if (changeS == "checkNo") /// ไม่ประสงค์ออกฝึก
            {
                if (ds.Tables[0].Rows[0]["confirm_status"].ToString() != "ไม่ประสงค์ออกฝึก")
                {
                    _update += "UPDATE Student SET confirm_status = 'ไม่ประสงค์ออกฝึก', semester_out = '"+sem+"' where student_id = '" + std.student_id + "';";
                }

            }


            ti.msgAlert = db.insert_update_delete(_update);
            return RedirectToAction("EditProfile", ti);
        }


        // GET First Time -------------------------------------------------------------------------------------
        public ActionResult FirstTime(string student_id)
        {
            ConnectDB db = new ConnectDB();
            var _login = "SELECT s.student_id, s.name_th, a.major FROM Student s,Account a where s.student_id = '" + student_id + "' AND a.user_id = '" + student_id + "' AND name_en is null ;";
            _login += "SELECT DISTINCT province FROM PrototypeAddress order by province ;";
            DataSet ds = new DataSet();
            ds = db.select(_login, "Student");

            Student std = new Student();
            if (ds.Tables["Student"].Rows.Count != 0)
            {
                std.student_id = ds.Tables["Student"].Rows[0][0].ToString();
                std.name_th = ds.Tables["Student"].Rows[0][1].ToString();
                std.major = ds.Tables["Student"].Rows[0][2].ToString();
            }

            List<string> ListProvince = new List<string>();
            int _count = ds.Tables["Student1"].Rows.Count;

            for (int i = 0; i < _count; i++)
            {
                ListProvince.Add(ds.Tables["Student1"].Rows[i]["province"].ToString());
            }

            ViewBag.province = ListProvince;
            
            return View(std);
        }
        [HttpPost]
        public ActionResult FirstTime(Student std)
        {
            
            ConnectDB select = new ConnectDB();
            var run_id_checkCom = "SELECT MAX(language_id) FROM Language ;SELECT MAX(address_id) FROM Address ;SELECT MAX(family_id) FROM Family ;SELECT MAX(history_id) FROM HistoryofEducation;";
            DataSet ds = new DataSet();
            ds = select.select(run_id_checkCom, "Company_Address_Contact");

            var _countlan = ds.Tables[0].Rows[0][0].ToString();
            var _countAdd = ds.Tables[1].Rows[0][0].ToString();
            var _countfam = ds.Tables[2].Rows[0][0].ToString();
            var _counthis = ds.Tables[3].Rows[0][0].ToString();

            if (_countlan == "") std.language_id = 1;
            else std.language_id = Int32.Parse(_countlan) + 1;
            if (_countAdd == "") std.address_id = 1;
            else std.address_id = Int32.Parse(_countAdd) + 1;
            if (_countfam == "") std.family_id = 1;
            else std.family_id = Int32.Parse(_countfam) + 1;
            if (_counthis == "") std.history_id = 1;
            else std.history_id = Int32.Parse(_counthis) + 1;

            DateTime date_birth = Convert.ToDateTime(std.date_birth);
            std.age = DateTime.Now.Year - date_birth.Year;
            if (date_birth.Month > DateTime.Now.Month) std.age = std.age - 1;
            else if (date_birth.Month == DateTime.Now.Month)
            {
                if (date_birth.Day > DateTime.Now.Day) std.age = std.age - 1;
            }

            var _update_insert = "UPDATE Student SET name_en = '" + std.name_en + "', study = '" + std.study + "', year = " + std.year + ", advisor = '" + std.advisor + "', gpa = " + std.gpa + ", "
                + "gpax = " + std.gpax + ", identification = '" + std.identification + "', at = '" + std.at + "', date = '" + std.date + "', telephone = '" + std.telephone + "', mobile_phone = '"
                + std.mobile_phone + "', fax = '" + std.fax + "', email = '" + std.email + "', race = '" + std.race + "', nationality = '" + std.nationality + "', religion = '" + std.religion + "', "
                + "date_birth = '" + std.date_birth + "', age = " + std.age + ", gender = '" + std.gender + "', height = " + std.height + ", weight = " + std.weight + ", disease = '" + std.disease
                + "', no_relative = " + std.no_relative + ", number = " + std.number + ", ability_spectify = '" + std.ability_spectify + "', experience = '" + std.experience + "', confirm_status = '"
                + std.confirm_status + "', address_id = " + std.address_id + " WHERE student_id = '" + std.student_id + "' ;";

            for (int i = 0; i < std.ad_number.Length; i++)
            {
                _update_insert += "INSERT INTO Address VALUES(" + std.address_id + ", '" + std.ad_prototype_id[i] + "', '" + std.student_id + "', '" + std.ad_who[i] + "', '" + std.ad_number[i] + "', '"
                    + std.ad_building[i] + "', '" + std.ad_floor[i] + "', '" + std.ad_lane[i] + "', '" + std.ad_road[i] + "','','') ;";
                std.address_id = std.address_id + 1;
            }

            std.la_number = 1;

            for (int i = 0; i < std.la_name.Length; i++)
            {
                if (std.la_name[i] == "") continue;
                _update_insert += "INSERT INTO Language VALUES(" + std.language_id + ", '" + std.student_id + "', " + std.la_number + ", '" + std.la_name[i] + "', '" + std.la_listening[i]
                    + "', '" + std.la_speaking[i] + "', '" + std.la_reading[i] + "', '" + std.la_writing[i] + "') ;";
                std.language_id = std.language_id + 1;
                std.la_number = std.la_number + 1;
            }

            //std.hi_number = 1;

            for (int i = 0; i < std.hi_name.Length; i++)
            {
                if (std.hi_name[i] == "") continue;
                _update_insert += "INSERT INTO HistoryofEducation VALUES(" + std.history_id + ", '" + std.student_id + "', " + (i+1) + ", '" + std.hi_name[i] + "', " + std.hi_year[i] + ", '" + std.hi_certificate[i]
                    + "', '" + std.hi_major[i] + "', " + std.hi_gpa[i] + ") ;";
                std.history_id = std.history_id + 1;
                //std.hi_number = std.hi_number + 1;
            }

            std.fa_number = 1;

            for (int i = 0; i < std.fa_name.Length; i++)
            {
                if (i == 0)
                {
                    _update_insert += "INSERT INTO Family VALUES(" + std.family_id + ", '" + std.student_id + "', " + std.fa_number + ", '" + std.fa_name[i] + "', '" + std.fa_relation[i] + "', '', '"
                        + std.fa_occupation[i] + "', '', '" + std.fa_telephone[i] + "', '" + std.fa_mobile_phone[i] + "', '" + std.fa_fax[i] + "', '" + std.fa_email[i] + "', '" + std.fa_place_work[i] + "') ;";
                    std.family_id = std.family_id + 1;
                    std.fa_number = std.fa_number + 1;
                }
                else if (i < 3)
                {
                    _update_insert += "INSERT INTO Family VALUES(" + std.family_id + ", '" + std.student_id + "', " + std.fa_number + ", '" + std.fa_name[i] + "', '" + std.fa_relation[i] + "', '" + std.fa_age[i - 1] + "', '"
                        + std.fa_occupation[i] + "', '', '', '" + std.fa_mobile_phone[i] + "', '', '', '') ;";
                    std.family_id = std.family_id + 1;
                    std.fa_number = std.fa_number + 1;
                }
                else
                {
                    _update_insert += "INSERT INTO Family VALUES(" + std.family_id + ", '" + std.student_id + "', " + std.fa_number + ", '" + std.fa_name[i] + "', 'พี่-น้อง', '" + std.fa_age[i - 1] + "', '"
                        + std.fa_occupation[i] + "', '" + std.fa_position[i - 3] + "', '', '', '', '', '" + std.fa_place_work[i - 2] + "') ;";
                    std.family_id = std.family_id + 1;
                    std.fa_number = std.fa_number + 1;
                }

            }

            //_update_insert += "insert into Document(student_id) values('" + std.student_id + "') ;";
            //_update_insert += "insert into Supervisor(student_id) values('" + std.student_id + "') ;";

            ConnectDB update = new ConnectDB();
            _ = update.insert_update_delete(_update_insert);

            return RedirectToAction("Index", new { user_id_title = std.student_id });

        }



        public ActionResult SelectCompany(TitileAll ti)
        {
           
            ConnectDB db1 = new ConnectDB();
            var check = "select confirm_status from Student where student_id = '" + ti.user_id_title + "';";
            DataSet ds1 = new DataSet();
            ds1 = db1.select(check, "company");
            if (ds1.Tables[0].Rows[0][0].ToString() == "" || ds1.Tables[0].Rows[0][0].ToString() == "รอการตอบรับ" || ds1.Tables[0].Rows[0][0].ToString() == "ปฏิเสธ" || ds1.Tables[0].Rows[0][0].ToString() == "รับแล้ว")
            {
                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ti.name_title;
                ViewBag.Page = "Select Company";
                ConnectDB db = new ConnectDB();
                var detail = "select c.company_id, c.name_th, c.name_en, r.major, r.require from Company c, ReceiveSemester r, Account a, Student s where c.company_id = r.company_id "
                    + "and r.major = a.major and a.user_id = '" + ti.user_id_title + "' and s.student_id = '" + ti.user_id_title + "' and r.semester_id = s.semester_out;";
                DataSet ds = new DataSet();
                ds = db.select(detail, "company");
                return View(ds);
            }
            else
            {
                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ti.name_title;
                ViewBag.Page = "Select Company";
                ViewData["check"] = "No";
                return View();
            }

        }



        public ActionResult RegisterCompany(string company_id, TitileAll ti)
        {
            ConnectDB db = new ConnectDB();
            var detail = "SELECT s.company_id, s.semester_id, s.semester_out, s.confirm_status FROM Student s WHERE student_id = '" + ti.user_id_title + "' ;";

            detail += "select r.require from ReceiveSemester r, Student s, Account a where s.student_id = '" + ti.user_id_title
                + "' and r.semester_id = s.semester_out and r.major = a.major and a.user_id = s.student_id and r.company_id = " + company_id + " ;";
            detail += "select count(*) count_id from Student where semester_out in (select semester_out from Student where student_id = '" + ti.user_id_title + "') and company_id = " + company_id +
                " and not confirm_status = 'ปฏิเสธ' ; ";
            DataSet ds = new DataSet();
            ds = db.select(detail, "Student");

            if (int.Parse(ds.Tables[2].Rows[0]["count_id"].ToString()) >= int.Parse(ds.Tables[1].Rows[0]["require"].ToString()))
            {
                return RedirectToAction("Index", new { user_id_title = ti.user_id_title, msgAlert = "สถานประกอบการนี้คนสมัครเต็มแล้ว" });
            }
            else if (ds.Tables["Student"].Rows[0]["confirm_status"].ToString() == "ไม่ประสงค์ออกฝึกงานในเทอมนี้" && ds.Tables["Student"].Rows[0]["semester_id"].ToString() == 
                ds.Tables["Student"].Rows[0]["semester_ount"].ToString())
            {
                return RedirectToAction("Index", new { user_id_title = ti.user_id_title, msgAlert = "คุณยังไม่ได้เลือกปีการศึกษาใหม่" });
            }
            else if (ds.Tables["Student"].Rows[0]["confirm_status"].ToString() == "รอการตอบรับ")
            {
                return RedirectToAction("Index", new { user_id_title = ti.user_id_title, msgAlert = "คุณได้เลือกสถานประกอบการแล้ว และยังรอการตอบรับ" }); 
            }
            else if (ds.Tables["Student"].Rows[0]["confirm_status"].ToString() == "รับแล้ว")
            {
                return RedirectToAction("Index", new { user_id_title = ti.user_id_title, msgAlert = "คุณได้เลือกสถานประกอบการแล้ว" });
            }
            else
            {
                var register = "UPDATE Student SET company_id = " + company_id + ", confirm_status = 'รอการตอบรับ' WHERE student_id = '" + ti.user_id_title + "' ;";
                _ = db.insert_update_delete(register);

                return RedirectToAction("Index", new { user_id_title = ti.user_id_title, msgAlert = "การสมัครสำเร็จ" });
            }  
        }



        public ActionResult DetailCompany(string company_id, TitileAll ti)
        {
            //company_id = "4";
            //student_id = "1160304620454";
            //ViewData["student_id"] = student_id;
            ViewData["user_id"] = ti.user_id_title;
            ViewData["name"] = ti.name_title;
            ViewBag.Page = "Detail Company";
            ConnectDB db = new ConnectDB();
            var detail = "SELECT * FROM Company WHERE company_id = "+ company_id +" ;";
            detail += "select* from Address a, PrototypeAddress p where a.address_id = " + company_id + " and a.prototype_id = p.prototype_id ;";
            DataSet ds = new DataSet();
            ds = db.select(detail, "company");
            string floor, lane, road, building;
            building = ds.Tables["company1"].Rows[0]["building"].ToString();
            if (building.Length <= 2 && building != "")
            {
                building = "หมู่ ";
            }
            else
            {
                building = "";
            }
            floor = ds.Tables["company1"].Rows[0]["floor"].ToString();
            if (floor.Length < 4 && floor != "")
            {
                floor = "ชั้น ";
            }
            else
            {
                floor = "";
            }
            lane = ds.Tables["company1"].Rows[0]["lane"].ToString();
            if (lane != "")
            {
                lane = "ซอย ";
            }
            road = ds.Tables["company1"].Rows[0]["road"].ToString();
            if (road != "")
            {
                road = "ถนน.";
            }
            ViewData["address"] = ds.Tables["company1"].Rows[0]["number"].ToString() +" "+ building + ds.Tables["company1"].Rows[0]["building"].ToString() + " " + floor +
                ds.Tables["company1"].Rows[0]["floor"].ToString() + " " + lane + ds.Tables["company1"].Rows[0]["lane"].ToString() + " " + road +
                ds.Tables["company1"].Rows[0]["road"].ToString() + " ต." + ds.Tables["company1"].Rows[0]["sub_area"].ToString() + " อ." +
                ds.Tables["company1"].Rows[0]["area"].ToString() + " จ." + ds.Tables["company1"].Rows[0]["province"].ToString() + " " +
                ds.Tables["company1"].Rows[0]["postal_code"].ToString();
            return View(ds);
        }



        public ActionResult More07_09(TitileAll ti)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + ti.user_id_title + "' ;";
            ds1 = db1.select(check, "check");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ti.name_title;
                ViewData["msgAlert"] = ti.msgAlert;
                ViewBag.Page = "Document 07, 09";
                ConnectDB db = new ConnectDB();
                DataSet ds = new DataSet();
                string _detail = "select * from Document where student_id = '" + ti.user_id_title + "' ;";
                _detail += "select * from Supervisor where student_id = '" + ti.user_id_title + "' ;";

                List<string> ListItems = new List<string>();
                _detail += "SELECT DISTINCT province FROM PrototypeAddress where not prototype_id = 0 order by province ;";

                _detail += "SELECT * FROM Address a, PrototypeAddress p WHERE a.student_company_id = '" + ti.user_id_title + "' and a.who = 'ปัจจุบัน' and p.prototype_id = a.prototype_id ;";
                _detail += "select distinct area from PrototypeAddress where province = (SELECT province FROM Address a, PrototypeAddress p WHERE a.student_company_id = '" + ti.user_id_title +
                    "' and a.who = 'ปัจจุบัน' and p.prototype_id = a.prototype_id) order by area ;";
                _detail += "select distinct sub_area from PrototypeAddress where area = (SELECT area FROM Address a, PrototypeAddress p WHERE a.student_company_id = '" + ti.user_id_title +
                    "' and a.who = 'ปัจจุบัน' and p.prototype_id = a.prototype_id) order by sub_area ;";
                ds = db.select(_detail, "Detail");

                int _count = ds.Tables["Detail2"].Rows.Count;
                for (int i = 0; i < _count; i++)
                {
                    ListItems.Add(ds.Tables["Detail2"].Rows[i]["province"].ToString());
                }
                ViewBag.province = ListItems;

                List<string> ListArea = new List<string>();
                int _countArea = ds.Tables["Detail4"].Rows.Count;
                for (int i = 0; i < _countArea; i++)
                {
                    ListArea.Add(ds.Tables["Detail4"].Rows[i]["area"].ToString());
                }


                List<string> ListSubArea = new List<string>();
                int _countSubArea = ds.Tables["Detail5"].Rows.Count;
                for (int i = 0; i < _countSubArea; i++)
                {
                    ListSubArea.Add(ds.Tables["Detail5"].Rows[i]["sub_area"].ToString());
                }

                if (ds.Tables["Detail3"].Rows.Count == 0)
                {
                    ds.Tables["Detail3"].Rows.Add();
                    ViewBag.area = "";
                    ViewBag.sub_area = "";
                }
                else
                {
                    ViewBag.area = ListArea;
                    ViewBag.sub_area = ListSubArea;
                }
                ViewData["prototype_id"] = ds.Tables["Detail3"].Rows[0]["prototype_id"].ToString();
                ViewData["target"] = "no";
                return View(ds);
            }
            else
                {
                    
                    return RedirectToAction("AlertNoCompany", "Student");
                }
            
        }
        [HttpPost]
        public ActionResult More07_09(SupDoc sd ,TitileAll ti)
        {

            ConnectDB db = new ConnectDB();
                string _check = "select * from Address where student_company_id = '" + ti.user_id_title + "' and who = 'ปัจจุบัน' ;";
                _check += "select max(address_id) from address ;";
                DataSet ds = new DataSet();
                ds = db.select(_check, "Check");
                //---------------------------------------------------------------------
                string _insert = "";
                if (ds.Tables["Check"].Rows.Count == 0)
                {
                    _insert += "insert into Address values(" + (int.Parse(ds.Tables["Check1"].Rows[0][0].ToString()) + 1) + ", " + sd.prototype_id + ", '" + ti.user_id_title + "', 'ปัจจุบัน', '" + sd.number + "', '"
                        + sd.building + "', '" + sd.floor + "', '" + sd.lane + "', '" + sd.road + "','',''); ";
                }
                else
                {
                    _insert += "update Address set prototype_id = " + sd.prototype_id + ", number = '" + sd.number + "', building = '" + sd.building + "', floor = '" + sd.floor +
                        "', lane = '" + sd.lane + "', road = '" + sd.road + "' where student_company_id = '" + ti.user_id_title + "' and who = 'ปัจจุบัน'; ";
                }
                _insert += "update Supervisor set name = '" + sd.name + "', position = '" + sd.position + "', department = '" + sd.department + "', phone = '" + sd.phone +
                    "', fax = '" + sd.fax + "', email = '" + sd.email + "' where student_id = '" + ti.user_id_title + "' ;";
                _insert += "update Document set title_en = '" + sd.title_en + "', title_th = '" + sd.title_th + "', title_detail = '" + sd.title_detail + "', job_position = '" + sd.job_position +
                    "', job_description = '" + sd.job_description + "', apartment = '" + sd.apartment + "', room = '" + sd.room + "', receive_doc = '" + sd.receive_doc + "' where student_id = '"
                    + ti.user_id_title + "' ;";
                ti.msgAlert = db.insert_update_delete(_insert);



            return RedirectToAction("More07_09", ti);
          
        }



        public ActionResult More08(TitileAll ti)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + ti.user_id_title + "' ;";
            ds1 = db1.select(check, "check");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ti.name_title;
                ViewData["msgAlert"] = ti.msgAlert;
                ViewBag.Page = "Document 08";
                ConnectDB db = new ConnectDB();
                string _check = "select * from Work where student_id = '" + ti.user_id_title + "' ;";
                DataSet ds = new DataSet();
                ds = db.select(_check, "Work");
                //ViewData["row_data"] = ds.Tables["Work"].Rows.Count;
                string all_week = "";
                List<int> _round = new List<int>();
                //int count_comma;
                string[] test;
                for (int i = 0; i < ds.Tables["Work"].Rows.Count; i++)
                {
                    //count_comma = 0;
                    all_week += ds.Tables["Work"].Rows[i]["week"].ToString() + ",";
                    /*
                    foreach (char c in ds.Tables["Work"].Rows[i]["week"].ToString())
                    {
                        if (c == ',') count_comma++;
                    }
                    _round.Add(ds.Tables["Work"].Rows[i]["week"].ToString().Length - count_comma);
                    */
                    //------------------------------------------------------------
                    test = ds.Tables["Work"].Rows[i]["week"].ToString().Split(',');
                    _round.Add(test.Length);
                }
                if (all_week != "") all_week = all_week.Remove(all_week.Length - 1);

                string[] list_week = all_week.Split(',');
                List<string> week = new List<string>();
                for (int i = 0; i < list_week.Length; i++)
                {
                    week.Add(list_week[i]);
                }
                //----------------------------------------------------------------
                for (int i = ds.Tables["Work"].Rows.Count; i < 12; i++)
                {
                    ds.Tables["Work"].Rows.Add();
                }

                DataTable dt = new DataTable("Round");
                dt.Columns.Add("round", typeof(int));
                foreach (int rnd in _round)
                {
                    DataRow row = dt.NewRow();
                    row["round"] = rnd;
                    dt.Rows.Add(row);
                }
                ds.Tables.Add(dt);

                DataTable dt1 = new DataTable("Week");
                dt1.Columns.Add("week", typeof(string));
                foreach (string wek in week)
                {
                    DataRow row = dt1.NewRow();
                    row["week"] = wek;
                    dt1.Rows.Add(row);
                }
                ds.Tables.Add(dt1);
                ViewData["target"] = "no";
                return View(ds);
            }
            else
            {
                return RedirectToAction("AlertNoCompany", "Student");
            }

        }
        [HttpPost]
        public ActionResult More08(Work08 wk, TitileAll ti)
        {

            ConnectDB db = new ConnectDB();
            var _check = "select max(works_id) from Work ;";
            _check += "select * from Work where student_id = '" + ti.user_id_title + "' ;";
            DataSet ds = new DataSet();
            ds = db.select(_check, "Work");
            if (ds.Tables["Work"].Rows[0][0].ToString() == "")
            {
                ds.Tables["Work"].Rows[0][0] = 0;
            }
            string _insert = "";
            string _week;
            int works_id = int.Parse(ds.Tables["Work"].Rows[0][0].ToString()) + 1;
            for (int i=0; i<12; i++)
            {
                if (wk.detail[i] == "")
                {
                    if (ds.Tables["Work1"].Rows.Count > i)
                    {
                        int _end = ds.Tables["Work1"].Rows.Count;
                        for (int k=i; k < _end; k++)
                        {
                            _insert += "delete from Work where student_id = '" + ti.user_id_title + "' and number = " + (k + 1) + " ;";
                        }
                    }
                    break;
                }
                else if ((i+1) <= ds.Tables["Work1"].Rows.Count)
                {
                    _week = "";
                    for (int j = i * 16; j < 16 * (i + 1); j++)
                    {
                        if (wk.week[j] == "on")
                        {
                            _week += (j + 1 - i * 16).ToString() + ",";
                        }
                    }
                    _week = _week.Remove(_week.Length - 1);
                    _insert += "update Work set detail = '" + wk.detail[i] + "', week = '" + _week + "' where student_id = '"+ti.user_id_title+"' and number = "+(i+1)+" ;";
                }
                else if ((i+1) > ds.Tables["Work1"].Rows.Count)
                {
                    _week = "";
                    for (int j = i * 16; j < 16 * (i + 1); j++)
                    {
                        if (wk.week[j] == "on")
                        {
                            _week += (j + 1 - i * 16).ToString() + ",";
                        }
                    }
                    _week = _week.Remove(_week.Length - 1);
                    _insert += "insert into Work values(" + works_id + ", '" + ti.user_id_title + "', " + (i + 1) + ", '" + wk.detail[i] + "', '"+_week+"') ;";
                    works_id++;
                }
            }
            ti.msgAlert = db.insert_update_delete(_insert);
            return RedirectToAction("More08", ti);
        } 


        public ActionResult No03(string student_id)
        {
            //student_id = "1160304620454";
            //ViewData["student_id"] = student_id;
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            Student std = new Student();
            string _select = "select c.name_en, c.name_th from Company c, Student s where s.student_id = '" + student_id + "' and s.company_id = c.company_id ;";
            _select += "select d.date_start, d.date_end from Student s, DefineSemester d where s.student_id = '" + student_id + "' and s.semester_out = d.semester_id ;";
            _select += "select * from Student s, Account a, Address ad, PrototypeAddress p where s.student_id = '" + student_id +
                "' and a.user_id = s.student_id and ad.address_id = s.address_id and p.prototype_id = ad.prototype_id ;";
            _select += "select * from Family where student_id = '" + student_id + "' ;";
            _select += "select * from Address ad, PrototypeAddress p where ad.student_company_id = '" + student_id + "' and p.prototype_id = ad.prototype_id and not ad.who = 'นักศึกษา' ;";
            _select += "select * from Family where student_id = '" + student_id + "' and relation = 'พี่-น้อง' ;";
            _select += "select * from HistoryofEducation where student_id = '" + student_id + "' ;";
            _select += "select * from Language where student_id = '" + student_id + "' ;";
            
            ds = db.select(_select,"Student");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว" || ds1.Tables["check"].Rows[0][0].ToString() == "รอการตอบรับ")
            {
                string company_name_en = ds.Tables["Student"].Rows[0]["name_en"].ToString();
                string company_name_th = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                if (company_name_th.Length < 5) company_name_th = company_name_en;
                string date_start = ds.Tables["Student1"].Rows[0]["date_start"].ToString();
                string date_end = ds.Tables["Student1"].Rows[0]["date_end"].ToString();

                std.name_en = ds.Tables["Student2"].Rows[0]["name_en"].ToString();
                std.name_th = ds.Tables["Student2"].Rows[0]["name_th"].ToString();
                //std.student_id = ds.Tables["Student2"].Rows[0]["student_id"].ToString();
                std.student_id = "";
                for (int i = 0; i < 12; i++)
                {
                    std.student_id += student_id[i].ToString();
                }
                std.student_id += "-";
                std.student_id += student_id[12].ToString();
                std.major = ds.Tables["Student2"].Rows[0]["major"].ToString();
                std.year = int.Parse(ds.Tables["Student2"].Rows[0]["year"].ToString());
                std.advisor = ds.Tables["Student2"].Rows[0]["advisor"].ToString();
                std.gpa = float.Parse(ds.Tables["Student2"].Rows[0]["gpa"].ToString());
                std.gpax = float.Parse(ds.Tables["Student2"].Rows[0]["gpax"].ToString());
                std.identification = ds.Tables["Student2"].Rows[0]["identification"].ToString();
                std.at = ds.Tables["Student2"].Rows[0]["at"].ToString();
                std.date = ds.Tables["Student2"].Rows[0]["date"].ToString();
                std.race = ds.Tables["Student2"].Rows[0]["race"].ToString();
                std.nationality = ds.Tables["Student2"].Rows[0]["nationality"].ToString();
                std.religion = ds.Tables["Student2"].Rows[0]["religion"].ToString();
                std.date_birth = ds.Tables["Student2"].Rows[0]["date_birth"].ToString();
                std.age = int.Parse(ds.Tables["Student2"].Rows[0]["age"].ToString());
                std.gender = ds.Tables["Student2"].Rows[0]["gender"].ToString();
                std.height = int.Parse(ds.Tables["Student2"].Rows[0]["height"].ToString());
                std.weight = int.Parse(ds.Tables["Student2"].Rows[0]["weight"].ToString());
                std.disease = ds.Tables["Student2"].Rows[0]["disease"].ToString();

                string floor, lane, road, building, address_name;
                building = ds.Tables["Student2"].Rows[0]["building"].ToString();
                if (building.Length <= 2 && building != "")
                {
                    building = "หมู่ ";
                }
                else
                {
                    building = "";
                }
                floor = ds.Tables["Student2"].Rows[0]["floor"].ToString();
                if (floor.Length < 4 && floor != "")
                {
                    floor = "ชั้น ";
                }
                else
                {
                    floor = "";
                }
                lane = ds.Tables["Student2"].Rows[0]["lane"].ToString();
                if (lane != "")
                {
                    lane = "ซอย ";
                }
                road = ds.Tables["Student2"].Rows[0]["road"].ToString();
                if (road != "")
                {
                    road = "ถนน.";
                }
                address_name = ds.Tables["Student2"].Rows[0]["number1"].ToString() + " " + building + ds.Tables["Student2"].Rows[0]["building"].ToString() + " " + floor +
                    ds.Tables["Student2"].Rows[0]["floor"].ToString() + " " + lane + ds.Tables["Student2"].Rows[0]["lane"].ToString() + " " + road +
                    ds.Tables["Student2"].Rows[0]["road"].ToString() + " ต." + ds.Tables["Student2"].Rows[0]["sub_area"].ToString() + " อ." +
                    ds.Tables["Student2"].Rows[0]["area"].ToString() + " จ." + ds.Tables["Student2"].Rows[0]["province"].ToString() + " " +
                    ds.Tables["Student2"].Rows[0]["postal_code"].ToString();

                std.telephone = ds.Tables["Student2"].Rows[0]["telephone"].ToString();
                std.mobile_phone = ds.Tables["Student2"].Rows[0]["mobile_phone"].ToString();
                std.fax = ds.Tables["Student2"].Rows[0]["fax"].ToString();
                std.email = ds.Tables["Student2"].Rows[0]["email"].ToString();
                std.number = int.Parse(ds.Tables["Student2"].Rows[0]["number"].ToString());
                std.no_relative = int.Parse(ds.Tables["Student2"].Rows[0]["no_relative"].ToString());
                std.ability_spectify = ds.Tables["Student2"].Rows[0]["ability_spectify"].ToString();
                std.experience = ds.Tables["Student2"].Rows[0]["experience"].ToString();

                int _countFam = ds.Tables["Student3"].Rows.Count;
                std.fa_name = new string[_countFam];
                std.fa_relation = new string[_countFam];
                std.fa_age = new int[_countFam];
                std.fa_occupation = new string[_countFam];
                std.fa_position = new string[_countFam];
                std.fa_telephone = new string[_countFam];
                std.fa_mobile_phone = new string[_countFam];
                std.fa_fax = new string[_countFam];
                std.fa_email = new string[_countFam];
                std.fa_place_work = new string[_countFam];
                for (int i = 0; i < _countFam; i++)
                {
                    std.fa_name[i] = ds.Tables["Student3"].Rows[i]["name"].ToString();
                    std.fa_relation[i] = ds.Tables["Student3"].Rows[i]["relation"].ToString();
                    std.fa_age[i] = int.Parse(ds.Tables["Student3"].Rows[i]["age"].ToString());
                    std.fa_occupation[i] = ds.Tables["Student3"].Rows[i]["occupation"].ToString();
                    std.fa_position[i] = ds.Tables["Student3"].Rows[i]["position"].ToString();
                    std.fa_telephone[i] = ds.Tables["Student3"].Rows[i]["telephone"].ToString();
                    std.fa_mobile_phone[i] = ds.Tables["Student3"].Rows[i]["mobile_phone"].ToString();
                    std.fa_fax[i] = ds.Tables["Student3"].Rows[i]["fax"].ToString();
                    std.fa_email[i] = ds.Tables["Student3"].Rows[i]["email"].ToString();
                    std.fa_place_work[i] = ds.Tables["Student3"].Rows[i]["place_work"].ToString();
                }

                string address_emergency;
                building = ds.Tables["Student4"].Rows[0]["building"].ToString();
                if (building.Length <= 2 && building != "")
                {
                    building = "หมู่ ";
                }
                else
                {
                    building = "";
                }
                floor = ds.Tables["Student4"].Rows[0]["floor"].ToString();
                if (floor.Length < 4 && floor != "")
                {
                    floor = "ชั้น ";
                }
                else
                {
                    floor = "";
                }
                lane = ds.Tables["Student4"].Rows[0]["lane"].ToString();
                if (lane != "")
                {
                    lane = "ซอย ";
                }
                road = ds.Tables["Student4"].Rows[0]["road"].ToString();
                if (road != "")
                {
                    road = "ถนน.";
                }
                address_emergency = ds.Tables["Student4"].Rows[0]["number"].ToString() + " " + building + ds.Tables["Student4"].Rows[0]["building"].ToString() + " " + floor +
                    ds.Tables["Student4"].Rows[0]["floor"].ToString() + " " + lane + ds.Tables["Student4"].Rows[0]["lane"].ToString() + " " + road +
                    ds.Tables["Student4"].Rows[0]["road"].ToString() + " ต." + ds.Tables["Student4"].Rows[0]["sub_area"].ToString() + " อ." +
                    ds.Tables["Student4"].Rows[0]["area"].ToString() + " จ." + ds.Tables["Student4"].Rows[0]["province"].ToString() + " " +
                    ds.Tables["Student4"].Rows[0]["postal_code"].ToString();
                string[] splitAdd = new string[2];
                string[] txtSplit = { "จ." };
                if (address_emergency.Length > 50)
                {
                    splitAdd = address_emergency.Split(txtSplit, StringSplitOptions.None);
                    splitAdd[1] = "จ." + splitAdd[1];
                }

                string address_dad = "", address_mom = "";
                for (int i = 1; i < ds.Tables["Student4"].Rows.Count; i++)
                {
                    if (ds.Tables["Student4"].Rows[i]["who"].ToString() == "บิดา")
                    {
                        building = ds.Tables["Student4"].Rows[i]["building"].ToString();
                        if (building.Length <= 2 && building != "")
                        {
                            building = "หมู่ ";
                        }
                        else
                        {
                            building = "";
                        }
                        floor = ds.Tables["Student4"].Rows[i]["floor"].ToString();
                        if (floor.Length < 4 && floor != "")
                        {
                            floor = "ชั้น ";
                        }
                        else
                        {
                            floor = "";
                        }
                        lane = ds.Tables["Student4"].Rows[i]["lane"].ToString();
                        if (lane != "")
                        {
                            lane = "ซอย ";
                        }
                        road = ds.Tables["Student4"].Rows[i]["road"].ToString();
                        if (road != "")
                        {
                            road = "ถนน.";
                        }
                        address_dad = ds.Tables["Student4"].Rows[i]["number"].ToString() + " " + building + ds.Tables["Student4"].Rows[i]["building"].ToString() + " " + floor +
                            ds.Tables["Student4"].Rows[i]["floor"].ToString() + " " + lane + ds.Tables["Student4"].Rows[i]["lane"].ToString() + " " + road +
                            ds.Tables["Student4"].Rows[i]["road"].ToString() + " ต." + ds.Tables["Student4"].Rows[i]["sub_area"].ToString() + " อ." +
                            ds.Tables["Student4"].Rows[i]["area"].ToString() + " จ." + ds.Tables["Student4"].Rows[i]["province"].ToString() + " " +
                            ds.Tables["Student4"].Rows[i]["postal_code"].ToString();
                    }
                    else if (ds.Tables["Student4"].Rows[i]["who"].ToString() == "มารดา")
                    {
                        building = ds.Tables["Student4"].Rows[i]["building"].ToString();
                        if (building.Length <= 2 && building != "")
                        {
                            building = "หมู่ ";
                        }
                        else
                        {
                            building = "";
                        }
                        floor = ds.Tables["Student4"].Rows[i]["floor"].ToString();
                        if (floor.Length < 4 && floor != "")
                        {
                            floor = "ชั้น ";
                        }
                        else
                        {
                            floor = "";
                        }
                        lane = ds.Tables["Student4"].Rows[i]["lane"].ToString();
                        if (lane != "")
                        {
                            lane = "ซอย ";
                        }
                        road = ds.Tables["Student4"].Rows[i]["road"].ToString();
                        if (road != "")
                        {
                            road = "ถนน.";
                        }
                        address_mom = ds.Tables["Student4"].Rows[i]["number"].ToString() + " " + building + ds.Tables["Student4"].Rows[i]["building"].ToString() + " " + floor +
                            ds.Tables["Student4"].Rows[i]["floor"].ToString() + " " + lane + ds.Tables["Student4"].Rows[i]["lane"].ToString() + " " + road +
                            ds.Tables["Student4"].Rows[i]["road"].ToString() + " ต." + ds.Tables["Student4"].Rows[i]["sub_area"].ToString() + " อ." +
                            ds.Tables["Student4"].Rows[i]["area"].ToString() + " จ." + ds.Tables["Student4"].Rows[i]["province"].ToString() + " " +
                            ds.Tables["Student4"].Rows[i]["postal_code"].ToString();
                    }
                }

                int _countRelative = ds.Tables["Student5"].Rows.Count;
                int _countEducation = 0;
                std.hi_name = new string[4] { "", "", "", "" };
                std.hi_year = new int[4];
                std.hi_certificate = new string[4] { "", "", "", "" };
                std.hi_major = new string[4] { "", "", "", "" };
                std.hi_gpa = new float[4];
                for (int i = 0; i < 4; i++)
                {
                    if ((_countEducation+1) > ds.Tables["Student6"].Rows.Count) break;
                    if (int.Parse(ds.Tables["Student6"].Rows[_countEducation]["number"].ToString()) != (i + 1)) continue;
                    std.hi_name[i] = ds.Tables["Student6"].Rows[_countEducation]["name"].ToString();
                    std.hi_year[i] = int.Parse(ds.Tables["Student6"].Rows[_countEducation]["year"].ToString());
                    std.hi_certificate[i] = ds.Tables["Student6"].Rows[_countEducation]["certificate"].ToString();
                    std.hi_major[i] = ds.Tables["Student6"].Rows[_countEducation]["major"].ToString();
                    std.hi_gpa[i] = float.Parse(ds.Tables["Student6"].Rows[_countEducation]["gpa"].ToString());
                    _countEducation++;
                }

                int _countLanguage = ds.Tables["Student7"].Rows.Count;
                std.la_name = new string[_countLanguage];
                std.la_listening = new string[_countLanguage];
                std.la_speaking = new string[_countLanguage];
                std.la_reading = new string[_countLanguage];
                std.la_writing = new string[_countLanguage];
                for (int i = 0; i < _countLanguage; i++)
                {
                    std.la_name[i] = ds.Tables["Student7"].Rows[i]["name"].ToString();
                    std.la_listening[i] = ds.Tables["Student7"].Rows[i]["listening"].ToString();
                    std.la_speaking[i] = ds.Tables["Student7"].Rows[i]["speaking"].ToString();
                    std.la_reading[i] = ds.Tables["Student7"].Rows[i]["reading"].ToString();
                    std.la_writing[i] = ds.Tables["Student7"].Rows[i]["writing"].ToString();
                }

                // PDF ------------------------------------------------------------------------------------------------------------------------
                string fileName;

                if (ds.Tables["Student2"].Rows[0]["semester_out"].ToString().Substring(5, 1) == "3")
                {
                    fileName = "no03-3.pdf";
                }
                else
                {
                    fileName = "no03.pdf";
                }

                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                string sourceFile = Path.Combine(sourcePath, fileName);
                string destFile = Path.Combine(targetPath, student_id + "_" + fileName);

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);


                System.IO.File.Copy(sourceFile, destFile, true);

                using (var reader = new PdfReader(oldFile))
                {
                    using (var fileStream = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                    {
                        var document = new Document(reader.GetPageSizeWithRotation(1));
                        var writer = PdfWriter.GetInstance(document, fileStream);

                        document.Open();

                        for (var i = 1; i <= reader.NumberOfPages; i++) // ใส่ข้อมูล 2 หน้า
                        {
                            BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                            // ใส่ข้อมูลหน้า 1
                            if (i == 1)
                            {
                                document.NewPage();

                                //BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);

                                cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, company_name_th, 250, 685, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, date_start, 292, 669, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, date_end, 458, 669, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.name_th, 190, 622, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.name_en, 230, 606, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.student_id, 234, 590, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.major, 488, 590, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.year.ToString(), 203, 574, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.advisor, 440, 574, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.gpa.ToString(), 237, 558, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.gpax.ToString(), 460, 558, 0);
                                //cb.EndText();
                                ///////////////////////////////////////////////////////// รหัสประชาชน
                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[0].ToString(), 295, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[1].ToString(), 318, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[2].ToString(), 330, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[3].ToString(), 342, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[4].ToString(), 353, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[5].ToString(), 377, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[6].ToString(), 389, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[7].ToString(), 401, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[8].ToString(), 412, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[9].ToString(), 424, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[10].ToString(), 447, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[11].ToString(), 459, 520, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.identification[12].ToString(), 483, 520, 0);
                                //cb.EndText();
                                /////////////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.at, 118, 495, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.date, 226, 495, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.race, 317, 495, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.nationality, 403, 495, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.religion, 500, 495, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.date_birth, 121, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.age.ToString(), 195, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.gender, 263, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.height.ToString(), 321, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.weight.ToString(), 385, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.disease, 505, 463, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, address_name, 131, 432, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.telephone, 103, 400, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.mobile_phone, 213, 400, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fax, 310, 400, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.email, 433, 400, 0);
                                //cb.EndText();
                                //////////////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_name[0], 230, 354, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_relation[0], 460, 354, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_occupation[0], 118, 322, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_place_work[0], 289, 322, 0);
                                //cb.EndText();

                                if (address_emergency.Length > 50)
                                {
                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitAdd[0], 380, 322, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, splitAdd[1], 395, 308, 0);
                                    //cb.EndText();
                                }
                                else
                                {
                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, address_emergency, 380, 322, 0);
                                    //cb.EndText();
                                }


                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_telephone[0], 102, 291, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_mobile_phone[0], 213, 291, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_fax[0], 310, 291, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.fa_email[0], 433, 291, 0);
                                //cb.EndText();
                                ////////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_name[1], 163, 244, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_age[1].ToString(), 304, 244, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_occupation[1], 450, 244, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, address_dad, 70, 228, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_mobile_phone[1], 465, 228, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_name[2], 163, 213, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_age[2].ToString(), 304, 213, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_occupation[2], 450, 213, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, address_mom, 70, 197, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_mobile_phone[2], 465, 197, 0);
                                //cb.EndText();
                                //////////////////////////////////////////////////////////////////// 

                                int line = 103;

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.no_relative.ToString(), 117, 180, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.number.ToString(), 236, 180, 0);
                                //cb.EndText();

                                for (int j = (_countFam - _countRelative); j < _countFam; j++)
                                {
                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_name[j], 127, line, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_age[j].ToString(), 217, line, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_occupation[j], 298, line, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_position[j], 385, line, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.fa_place_work[j], 476, line, 0);
                                    //cb.EndText();

                                    line = line - 16; // เว้นบรรทั้ด
                                }
                                cb.EndText();
                                cb.AddTemplate(importedPage, 0, 0);
                            }

                            // ใส่ข้อมูลหน้า 2
                            if (i == 2)
                            {
                                document.NewPage();

                                //BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);

                                /////////////////////////////////////////////////////// education
                                int many_education_line = 695;
                                cb.BeginText();
                                for (int j = 0; j < 4; j++)
                                {
                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.hi_name[j], 215, many_education_line, 0);
                                    //cb.EndText();

                                    if (std.hi_year[j] != 0)
                                    {
                                        //cb.BeginText();
                                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.hi_year[j].ToString(), 324, many_education_line, 0);
                                        //cb.EndText();
                                    }
                                    else
                                    {
                                        //cb.BeginText();
                                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "", 324, many_education_line, 0);
                                        //cb.EndText();
                                    }

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.hi_certificate[j], 389, many_education_line, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.hi_major[j], 450, many_education_line, 0);
                                    //cb.EndText();

                                    if (std.hi_gpa[j] != 0)
                                    {
                                        //cb.BeginText();
                                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.hi_gpa[j].ToString(), 521, many_education_line, 0);
                                        //cb.EndText();
                                    }
                                    else
                                    {
                                        //cb.BeginText();
                                        cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "", 521, many_education_line, 0);
                                        //cb.EndText();
                                    }

                                    many_education_line = (many_education_line - 30) + 1;
                                }

                                ///////////////////////////////////////////////////////////// เครื่องหมายถูก ✔ ** ถ้ามีหลายภาษา -30 บรรทัด

                                if (_countLanguage != 0)
                                {
                                    Stream inputImageStream = new FileStream(Server.MapPath("~/Images/correct_symbol.png"), FileMode.Open, FileAccess.Read, FileShare.Read);
                                    Image image = Image.GetInstance(inputImageStream);
                                    image.ScaleAbsolute(20, 20);
                                    int _correctY = 520;
                                    int _search;
                                    for (int j = 0; j < _countLanguage; j++)
                                    {
                                        /*
                                        std.la_name[i] = ds.Tables["Student6"].Rows[i]["name"].ToString();
                                        std.la_listening[i] = ds.Tables["Student6"].Rows[i]["listening"].ToString();
                                        std.la_speaking[i] = ds.Tables["Student6"].Rows[i]["speaking"].ToString();
                                        std.la_reading[i] = ds.Tables["Student6"].Rows[i]["reading"].ToString();
                                        std.la_writing[i] = ds.Tables["Student6"].Rows[i]["writing"].ToString();
                                        */
                                        _search = std.la_name[j].IndexOf("จีน");
                                        if (std.la_name[j] == "ภาษาอังกฤษ" || _search != -1)
                                        {
                                            if (std.la_listening[j] == "ดี")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// liste good
                                                image.SetAbsolutePosition(152, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else if (std.la_listening[j] == "ปานกลาง")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// liste fair
                                                image.SetAbsolutePosition(183, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// liste poor
                                                image.SetAbsolutePosition(213, _correctY);
                                                cb.AddImage(image);
                                            }

                                            //////////////////////////////////////////////////////////////////

                                            if (std.la_speaking[j] == "ดี")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// speak good
                                                image.SetAbsolutePosition(263, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else if (std.la_speaking[j] == "ปานกลาง")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// speak fair
                                                image.SetAbsolutePosition(295, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// speak poor
                                                image.SetAbsolutePosition(326, _correctY);
                                                cb.AddImage(image);
                                            }


                                            //////////////////////////////////////////////////////////////
                                            ///
                                            if (std.la_reading[j] == "ดี")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// reading good
                                                image.SetAbsolutePosition(371, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else if (std.la_reading[j] == "ปานกลาง")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// reading fair
                                                image.SetAbsolutePosition(402, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// reading poor
                                                image.SetAbsolutePosition(434, _correctY);
                                                cb.AddImage(image);
                                            }

                                            //////////////////////////////////////////////////////////////

                                            if (std.la_writing[j] == "ดี")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// writing good
                                                image.SetAbsolutePosition(469, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else if (std.la_writing[j] == "ปานกลาง")
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// writing fair
                                                image.SetAbsolutePosition(501, _correctY);
                                                cb.AddImage(image);
                                            }
                                            else
                                            {
                                                //image.ScaleAbsolute(20, 20);  //// writing poor
                                                image.SetAbsolutePosition(532, _correctY);
                                                cb.AddImage(image);
                                            }
                                        }
                                        _correctY -= 32;
                                    }
                                }

                                //////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.ability_spectify, 150, 414, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, std.experience, 175, 381, 0);
                                cb.EndText();

                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, std.name_th, 383, 295, 0);

                                cb.AddTemplate(importedPage, 0, 0);
                            }
                        }
                        document.Close();
                        writer.Close();
                    }
                }

                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                System.IO.File.Delete(newFile);
                return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("AlertNoCompany", "Student");
            }
        }
        public ActionResult No06(string student_id)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            string _select = "select * from Student s, Account a where s.student_id = '"+student_id+"' and s.student_id = a.user_id ;";
            _select += "select* from DefineSemester where semester_id = (select semester_out from Student where student_id = '" + student_id + "') ;";
            _select += "select name_th, name_en from Company where company_id = (select company_id from Student where student_id = '" + student_id + "') ;";
            _select += "select* from Contact where student_company_id like (select company_id from Student where student_id = '" + student_id + "') ;";
            _select += "select  t.name  from Student s inner join `match` m on s.student_id = m.student_id inner join TeachersandAuthorities t on m.teacher_id = t.teacher_id where s.student_id = '" + student_id + "' ;";
            ds = db.select(_select, "Student");
            string std_id = student_id.Substring(0, 12);
            string std_id1 = student_id.Substring(12, 1);
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                if (ds.Tables["Student2"].Rows[0]["name_th"].ToString().Length < 5)
                {
                    ds.Tables["Student2"].Rows[0]["name_th"] = ds.Tables["Student2"].Rows[0]["name_en"];
                }
                DateTime _dateStart = Convert.ToDateTime(ds.Tables["Student1"].Rows[0]["date_start"].ToString());
                DateTime _dateEnd = Convert.ToDateTime(ds.Tables["Student1"].Rows[0]["date_end"].ToString());
                Month _month = new Month();
                //*****************************************************************************
                string fileName = "no06.pdf";
                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                string sourceFile = Path.Combine(sourcePath, fileName);
                string destFile = Path.Combine(targetPath, student_id + "_" + fileName);

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);

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
                BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 12);

                // write the text in the pdf content
                string text;

                cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 241, 668, 0);
                //cb.EndText();

                //cb.BeginText();
                text = std_id + "-" + std_id1;
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 399, 668, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["age"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 492, 668, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["year"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 155, 647, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["major"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 306, 647, 0);
                //cb.EndText();

                //cb.BeginText();
                text = "วิศวกรรมศาสตร์";
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 475, 647, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student2"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 430, 626, 0);
                //cb.EndText();

                if (ds.Tables["Student3"].Rows.Count == 1)
                {
                    //cb.BeginText();
                    text = ds.Tables["Student3"].Rows[0]["name"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 131, 604, 0);
                    //cb.EndText();

                    //cb.BeginText();
                    text = ds.Tables["Student3"].Rows[0]["position"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 305, 604, 0);
                    //cb.EndText();
                }
                else if (ds.Tables["Student3"].Rows.Count == 2)
                {
                    //cb.BeginText();
                    text = ds.Tables["Student3"].Rows[1]["name"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 131, 604, 0);
                    //cb.EndText();

                    //cb.BeginText();
                    text = ds.Tables["Student3"].Rows[1]["position"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 305, 604, 0);
                    //cb.EndText();
                }


                //cb.BeginText();
                text = _dateStart.Day.ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 148, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = _month.get_month(_dateStart.Month);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 203, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = (_dateStart.Year + 543).ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 264, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = _dateEnd.Day.ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 321, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = _month.get_month(_dateEnd.Month);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 392, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = (_dateEnd.Year + 543).ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 471, 583, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student2"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 219, 383, 0);
                //EndText();

                //cb.BeginText();
                text = "วิศวกรรมศาสตร์";
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 131, 305, 0);
                cb.EndText();
                if (ds.Tables["Student4"].Rows.Count > 0)
                {
                    text = ds.Tables["Student4"].Rows[0]["name"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 140, 118, 0);
                }

                
                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 140, 214, 0);
                // create the new page and add it to the pdf
                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                cb.AddTemplate(page, 0, 0);

                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();

                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                System.IO.File.Delete(newFile);
                return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("AlertNoCompany", "Student");
            }
        }
        public ActionResult No07(string student_id)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            string _select = "select s.name_th, a.major from Student s, Account a where s.student_id = '"+student_id+"' and s.student_id = a.user_id ;";
            _select += "select* from Company c, Address a, PrototypeAddress p where company_id = (select company_id from Student s where s.student_id = '" + student_id + "') and c.address_id = a.address_id and a.prototype_id = p.prototype_id ;";
            _select += "select* from Contact where student_company_id like (select company_id from Student s where s.student_id = '" + student_id + "') ;";
            _select += "select* from Supervisor where student_id = '" + student_id + "' ;";
            _select += "select* from Family where student_id = '" + student_id + "' and number = 1 ;";
            _select += "select* from Address a, PrototypeAddress p where student_company_id = '" + student_id + "' and a.who in ('กรณีฉุกเฉิน', 'ปัจจุบัน') and a.prototype_id = p.prototype_id ;";
            _select += "select* from Document where student_id = '" + student_id + "' ;";
            ds = db.select(_select,"Student");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                if (ds.Tables["Student5"].Rows.Count == 1)
                {
                    return RedirectToAction("Index", new { user_id_title = student_id });
                }

                string std_id = student_id.Substring(0, 12);
                string std_id1 = student_id.Substring(12, 1);
                //****************************************************************************************
                string fileName = "no07.pdf";
                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, student_id + "_" + fileName);
                System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);

                string maps_pic = null;

                //////////////////////////////////////////////////////////////////////////////

                using (var reader = new PdfReader(oldFile))
                {
                    using (var fileStream = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                    {
                        var document = new Document(reader.GetPageSizeWithRotation(1));
                        var writer = PdfWriter.GetInstance(document, fileStream);

                        document.Open();

                        for (var i = 1; i <= reader.NumberOfPages; i++) // ใส่ข้อมูล 2 หน้า
                        {
                            BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

                            if (i == 1) // ใส่ข้อมูลหน้า 1
                            {
                                string text;
                                document.NewPage();

                                //BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);

                                cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["name_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 120, 660, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["name_en"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 125, 646, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["number"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 109, 632, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["lane"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 188, 632, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["road"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 317, 632, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["sub_area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 482, 632, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 170, 618, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["province"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 323, 618, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["postal_code"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 486, 618, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["phone"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 145, 605, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["fax"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 273, 605, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["email"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 441, 605, 0);
                                //cb.EndText();
                                ///////////////////////////////////////////////////////////////////////
                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["name"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 255, 550, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["position"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 456, 550, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["phone"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 165, 537, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["fax"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 310, 537, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["email"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 468, 539, 0);
                                //cb.EndText();
                                ////////////////////////////////////////////////////////////////////////
                                //cb.BeginText();
                                if (ds.Tables["Student2"].Rows.Count == 2)
                                {
                                    text = ds.Tables["Student2"].Rows[1]["name"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 140, 496, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    text = ds.Tables["Student2"].Rows[1]["position"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 213, 483, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    text = ds.Tables["Student2"].Rows[1]["department"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 445, 483, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    text = ds.Tables["Student2"].Rows[1]["phone"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 165, 471, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    text = ds.Tables["Student2"].Rows[1]["fax"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 310, 471, 0);
                                    //cb.EndText();

                                    //cb.BeginText();
                                    text = ds.Tables["Student2"].Rows[1]["email"].ToString();
                                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 468, 473, 0);
                                    //cb.EndText();
                                }

                                //////////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["name"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 140, 392, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["position"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 213, 379, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["department"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 445, 379, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["phone"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 165, 367, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["fax"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 310, 367, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["email"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 468, 369, 0);
                                //cb.EndText();
                                ////////////////////////////////////////////////////////////////////
                                //cb.BeginText();
                                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 203, 302, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = std_id + "-" + std_id1;
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 440, 302, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student"].Rows[0]["major"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 200, 288, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = "วิศวกรรมศาสตร์";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 441, 288, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student6"].Rows[0]["job_position"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 220, 268, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student6"].Rows[0]["job_description"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 230, 254, 0);
                                //cb.EndText();
                                //////////////////////////////////////////////////////////////////
                                //cb.BeginText();
                                text = ds.Tables["Student4"].Rows[0]["name"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 140, 157, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["number"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 102, 143, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["lane"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 163, 143, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["road"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 241, 143, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["sub_area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 350, 143, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 483, 143, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["province"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 128, 129, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[0]["postal_code"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 227, 129, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student4"].Rows[0]["mobile_phone"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 305, 129, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student4"].Rows[0]["fax"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 384, 129, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student4"].Rows[0]["email"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 486, 129, 0);
                                cb.EndText();

                                cb.AddTemplate(importedPage, 0, 0);
                            }

                            if (i == 2) // ใส่ข้อมูลหน้า 2
                            {
                                string text;
                                document.NewPage();

                                //BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);
                                cb.AddTemplate(importedPage, 0, 0);
                                cb.BeginText();
                                text = ds.Tables["Student6"].Rows[0]["apartment"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 268, 735, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student6"].Rows[0]["room"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 482, 735, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["number"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 100, 716, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["lane"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 210, 716, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["road"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 364, 716, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["building"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 495, 716, 0);
                                //cb.EndText();
                                ///////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["sub_area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 147, 696, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["area"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 308, 696, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["province"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 471, 696, 0);
                                //cb.EndText();
                                /////////////////////////////////////////////////////////////////////

                                //cb.BeginText();
                                text = ds.Tables["Student5"].Rows[1]["postal_code"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 153, 676, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["phone"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 313, 676, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["fax"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 477, 676, 0);
                                cb.EndText();

                                Stream inputImageStreamCorrect = new FileStream(Server.MapPath("~/images/correct_symbol.png"), FileMode.Open, FileAccess.Read, FileShare.Read);
                                Image imageCorrect = Image.GetInstance(inputImageStreamCorrect);
                                if (ds.Tables["Student6"].Rows[0]["receive_doc"].ToString() == "ไม่รับ")
                                {
                                    imageCorrect.ScaleAbsolute(9, 9);
                                    imageCorrect.SetAbsolutePosition(58, 640);
                                    cb.AddImage(imageCorrect);
                                }
                                else
                                {
                                    imageCorrect.ScaleAbsolute(9, 9);
                                    imageCorrect.SetAbsolutePosition(58, 625);
                                    cb.AddImage(imageCorrect);
                                    if (ds.Tables["Student6"].Rows[0]["receive_doc"].ToString() == "รับส่งไปยัง")
                                    {

                                        imageCorrect.ScaleAbsolute(9, 9);
                                        imageCorrect.SetAbsolutePosition(277, 625);
                                        cb.AddImage(imageCorrect);


                                    }
                                    else if (ds.Tables["Student6"].Rows[0]["receive_doc"].ToString() == "รับส่งไปยังที่พัก")
                                    {
                                        imageCorrect.ScaleAbsolute(9, 9);  //// ภาคพิเศษ
                                        imageCorrect.SetAbsolutePosition(232, 625);
                                        cb.AddImage(imageCorrect);

                                    }
                                }
                                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 150, 264, 0);



                                ///////////////////////////////////////////////////////////// รูปแผนที่

                                

                                if (st_map_pic != null)
                                {
                                    Stream inputImageStream = new FileStream(Server.MapPath("~/Images/maps/" + student_id + "-map.jpg"), FileMode.Open, FileAccess.Read, FileShare.Read);
                                    Image image = Image.GetInstance(inputImageStream);
                                    
                                    //image.ScaleAbsoluteHeight(209f);
                                    //image.ScaleAbsoluteWidth(321f);
                                    image.ScaleAbsolute(321f, 209f);
                                    image.SetAbsolutePosition(142, 332);
                                    cb.AddImage(image);
                                    inputImageStream.Close();
               
                                }


                            }
                        }
                        document.Close();
                        writer.Close();
                    }
                }
                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                //string picFile = Server.MapPath("~/Images/maps/" + student_id + "-map.jpg");
                //System.IO.File.Delete(picFile);
                System.IO.File.Delete(newFile);

                    return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("ALertNoCompany", "Student");
            }
        }
        public ActionResult No08(string student_id)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            Student std = new Student();
            string _select = "select s.student_id, s.name_th, a.major, s.semester_out, c.name_th, c.name_en from Student s, Account a, Company c where s.student_id = '" 
                + student_id + "' and a.user_id = s.student_id and c.company_id = s.company_id ;";
            _select += "select* from Work where student_id = '" + student_id + "' ;";
            ds = db.select(_select, "Student");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                if (ds.Tables["Student1"].Rows.Count == 0)
                {
                    return RedirectToAction("Index", new { user_id_title = student_id });
                }
                if (ds.Tables["Student"].Rows[0]["name_th1"].ToString().Length < 5)
                {
                    ds.Tables["Student"].Rows[0]["name_th1"] = ds.Tables["Student"].Rows[0]["name_en"];
                }
                string[] semester = ds.Tables["Student"].Rows[0]["semester_out"].ToString().Split('-');
                string std_id = ds.Tables["Student"].Rows[0]["student_id"].ToString().Substring(0, 12);
                string std_id1 = ds.Tables["Student"].Rows[0]["student_id"].ToString().Substring(12, 1);

                string all_week = "";
                List<int> _round = new List<int>();    //----------------------------------- ใช้ตรงนี้
                //int count_comma;
                string[] test;
                for (int i = 0; i < ds.Tables["Student1"].Rows.Count; i++)
                {
                    
                    //count_comma = 0;
                    all_week += ds.Tables["Student1"].Rows[i]["week"].ToString() + ",";
                    /*
                    foreach (char c in ds.Tables["Student1"].Rows[i]["week"].ToString())
                    {
                        if (c == ',') count_comma++;
                    }
                    _round.Add(ds.Tables["Student1"].Rows[i]["week"].ToString().Length - count_comma);
                    */
                    //-------------------------------------------------------------
                    test = ds.Tables["Student1"].Rows[i]["week"].ToString().Split(',');
                    _round.Add(test.Length);
                }
                all_week = all_week.Remove(all_week.Length - 1);
                string[] list_week = all_week.Split(',');
                List<int> week = new List<int>();   //----------------------------------- ใช้ตรงนี้
                for (int i = 0; i < list_week.Length; i++)
                {
                    week.Add(int.Parse(list_week[i]));
                }
                //-------------------------------------------------------------------------------------------
                string fileName = "no08.pdf";
                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, student_id + "_" + fileName);
                System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);

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
                BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 12);

                // write the text in the pdf content
                string text;
                PdfImportedPage page = writer.GetImportedPage(reader, 1);
                cb.AddTemplate(page, 0, 0);

                cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 220, 698, 0);
                //cb.EndText();

                //cb.BeginText();
                text = std_id + "-" + std_id1;
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 450, 698, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["major"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 202, 678, 0);
                //cb.EndText();

                //cb.BeginText();
                text = "วิศวกรรมศาสตร์";
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 445, 678, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["name_th1"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 210, 658, 0);
                //cb.EndText();

                //cb.BeginText();
                text = semester[1];
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 165, 638, 0);
                //cb.EndText();

                //cb.BeginText();
                text = semester[0];
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 278, 638, 0);

              
                cb.EndText();

              

                ///////////////////////////////////////////////////////////// เครื่องหมายถูก ✔ 

                Stream inputImageStream = new FileStream(Server.MapPath("~/images/correct_symbol.png"), FileMode.Open, FileAccess.Read, FileShare.Read);
                Image image = Image.GetInstance(inputImageStream);
                image.ScaleAbsolute(10, 10);

                float line_x;
                float line = 540;
                int total_week = 0;
                for (int m = 0; m < ds.Tables["Student1"].Rows.Count; m++)
                {
                    cb.BeginText();
                    text = (m + 1).ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 56, line, 0);
                    //cb.EndText();

                    //cb.BeginText();
                    text = ds.Tables["Student1"].Rows[m]["detail"].ToString();
                    cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 85, line, 0);
                    cb.EndText();
                    /*
                    line_x = 348;
                    for (int i = 0; i < 16; i++)
                    {
                        image.SetAbsolutePosition(line_x, line);
                        cb.AddImage(image);
                        line_x = line_x + (float)13.5;
                    }
                    */
                    for (int i = 0; i < _round[m]; i++)
                    {
                        line_x = 348;
                        line_x = line_x + (float)13.5 * week[i + total_week] - (float)13.5;
                        image.SetAbsolutePosition(line_x, line);
                        cb.AddImage(image);
                    }
                    line = line - (float)20.5;
                    total_week = total_week + _round[m];
                }

                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 165, 220, 0);
                // create the new page and add it to the pdf
               
                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();

                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                System.IO.File.Delete(newFile);
                return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("ALertNoCompany", "Student");
            }
        }
        public ActionResult No09(string student_id)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            
            ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            string _select = "select s.name_th, a.major, s.semester_out, s. study from Student s, Account a where s.student_id = '"+student_id+"' and s.student_id = a.user_id ;";
            _select += "select* from Company c, Address a, PrototypeAddress p where company_id = (select company_id from Student where student_id = '" + student_id + 
                "') and c.address_id = a.address_id and a.prototype_id = p.prototype_id ;";
            _select += "select* from Document where student_id = '" + student_id + "' ;";
            ds = db.select(_select, "Student");
            string std_id = student_id.Substring(0, 12);
            string std_id1 = student_id.Substring(12, 1);
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                if (ds.Tables["Student1"].Rows[0]["name_th"].ToString().Length < 5)
                {
                    ds.Tables["Student1"].Rows[0]["name_th"] = ds.Tables["Student1"].Rows[0]["name_en"];
                }
                //----------------------------------------
                string fileName = "no09.pdf";
                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, student_id + "_" + fileName);
                System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);

                // open the reader
                PdfReader reader = new PdfReader(oldFile);
                Rectangle size = reader.GetPageSizeWithRotation(1);
                Document document = new Document(size);

                // open the writer
                FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write);
                PdfWriter writer = PdfWriter.GetInstance(document, fs);

                // create the new page and add it to the pdf
                PdfImportedPage page = writer.GetImportedPage(reader, 1);

                document.Open();

                // the pdf content
                PdfContentByte cb = writer.DirectContent;

                cb.AddTemplate(page, 0, 0);

                // select the font properties
                BaseFont bf = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                cb.SetColorFill(BaseColor.BLACK);
                cb.SetFontAndSize(bf, 12);

                // write the text in the pdf content
                string text;

                cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 225, 473, 0);
                //cb.EndText();

                //cb.BeginText();
                text = std_id + "-" + std_id1;
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 455, 473, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["major"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 191, 457, 0);
                //cb.EndText();

                //cb.BeginText();
                text = "วิศวกรรมศาสตร์";
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 381, 457, 0);
                //cb.EndText();

                ///////////////////////////////////////////////////////////// เครื่องหมายถูก ✔ 

                Stream inputImageStream = new FileStream(Server.MapPath("~/images/correct_symbol.png"), FileMode.Open, FileAccess.Read, FileShare.Read);
                Image image = Image.GetInstance(inputImageStream);

                if (ds.Tables["Student"].Rows[0]["study"].ToString() == "ภาคปกติ")
                {
                    image.ScaleAbsolute(9, 9);  //// ภาคปกติ
                    image.SetAbsolutePosition(455, 454);
                    cb.AddImage(image);
                }
                else
                {
                    image.ScaleAbsolute(9, 9);  //// ภาคพิเศษ
                    image.SetAbsolutePosition(506, 454);
                    cb.AddImage(image);
                }

                /////////////////////////////////////////////////////////////////

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 280, 441, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["semester_out"].ToString().Substring(4, 1);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 437, 441, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student"].Rows[0]["semester_out"].ToString().Substring(0, 4);
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 519, 441, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["number"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 99, 425, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["lane"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 173, 425, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["road"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 264, 425, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["sub_area"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 386, 425, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["area"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 504, 425, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["province"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 110, 409, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["postal_code"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 210, 409, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["phone"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 296, 409, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["fax"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 392, 409, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student1"].Rows[0]["email"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 500, 409, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student2"].Rows[0]["title_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 125, 341, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student2"].Rows[0]["title_en"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 133, 325, 0);
                //cb.EndText();

                //cb.BeginText();
                text = ds.Tables["Student2"].Rows[0]["title_detail"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 85, 255, 0);
                cb.EndText();

                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 169, 130, 0);

                // close the streams and voilá the file should be changed :)
                document.Close();
                fs.Close();
                writer.Close();
                reader.Close();

                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                System.IO.File.Delete(newFile);
                return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("AlertNoCompany", "Student");
            }
        }
        public ActionResult No15(string student_id)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + student_id + "' ;";
            ds1 = db1.select(check, "check");
            
             ConnectDB db = new ConnectDB();
            DataSet ds = new DataSet();
            string _select = "select s.name_th, a.major, s.semester_out, s. study from Student s, Account a where s.student_id = '"+ student_id + "' and s.student_id = a.user_id ;";
            _select += "select* from Supervisor where student_id = '" + student_id + "' ;";
            _select += "select* from Document where student_id = '" + student_id + "' ;";
            _select += "select name_en, name_th from Company where company_id = (select company_id from Student where student_id = '" + student_id + "') ;";
            ds = db.select(_select, "Student");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")  ///////////////////////////////////////////////////////////////////////
            {
                if (ds.Tables["Student3"].Rows[0]["name_th"].ToString().Length < 5)
                {
                    ds.Tables["Student3"].Rows[0]["name_th"] = ds.Tables["Student1"].Rows[0]["name_en"];
                }
                //---------------------------------------------------------------------------
                string fileName = "no15.pdf";
                string sourcePath = Server.MapPath("~/prototypeNo");
                string targetPath = Server.MapPath("~/noStore/");

                // Use Path class to manipulate file and directory paths.
                string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
                string destFile = System.IO.Path.Combine(targetPath, student_id + "_" + fileName);
                System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

                string oldFile = Server.MapPath("~/prototypeNo/" + fileName);
                string newFile = Server.MapPath("~/noStore/" + student_id + "_" + fileName);

                using (var reader = new PdfReader(oldFile))
                {
                    using (var fileStream = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                    {
                        var document = new Document(reader.GetPageSizeWithRotation(1));
                        var writer = PdfWriter.GetInstance(document, fileStream);

                        document.Open();

                        for (var i = 1; i <= reader.NumberOfPages; i++) // ใส่ข้อมูล 2 หน้า
                        {
                            // ใส่ข้อมูลหน้า 1
                            if (i == 1)
                            {
                                string text;
                                document.NewPage();

                                BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);

                                cb.BeginText();
                                text = ds.Tables["Student"].Rows[0]["name_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 230, 542, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = student_id.Substring(0, 12) + "-" + student_id.Substring(12, 1);
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 460, 542, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student"].Rows[0]["major"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 220, 523, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = "วิศวกรรมศาสตร์";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 450, 523, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student3"].Rows[0]["name_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 160, 505, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["name"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 163, 486, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["position"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 202, 468, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student1"].Rows[0]["department"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 443, 468, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["title_th"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 140, 411, 0);
                                //cb.EndText();

                                //cb.BeginText();
                                text = ds.Tables["Student2"].Rows[0]["title_en"].ToString();
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 163, 394, 0);
                                cb.EndText();

                                cb.AddTemplate(importedPage, 0, 0);
                            }

                            if (i == 2)
                            {
                                document.NewPage();

                                BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.RED);
                                cb.SetFontAndSize(baseFont, 12);

                                cb.AddTemplate(importedPage, 0, 0);
                            }
                        }

                        document.Close();
                        writer.Close();
                    }
                }
                string ReportURL = Server.MapPath("~/noStore/" + student_id + "_" + fileName);
                byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);
                System.IO.File.Delete(newFile);
                return File(FileBytes, "application/pdf");
            }
            else
            {
                return RedirectToAction("AlertNoCompany", "Student");
            }
        }

        public ActionResult AlertNoCompany() //////มึงไม่ได้เลือกบริษัท///////////
        {
            return View();
        }

        public ActionResult uploadmap(TitileAll ti)
        {
            ConnectDB db1 = new ConnectDB();
            DataSet ds1 = new DataSet();
            string check = "select confirm_status from Student where student_id = '" + ti.user_id_title + "' ;";
            ds1 = db1.select(check, "check");
            if (ds1.Tables["check"].Rows[0][0].ToString() == "รับแล้ว")
            {
                
                ViewData["user_id"] = ti.user_id_title;
                ViewData["name"] = ti.name_title;
                return View();
            }
            else { 
                return RedirectToAction("AlertNoCompany", "Student");
        }
    }

        [HttpPost]
        public ActionResult upload_map(HttpPostedFileBase file, TitileAll ti)
        {
            if (file != null)
            {


                    string ImageName = ti.user_id_title + "-map.jpg";

                    st_map_pic = ImageName;

                    string physicalPath = Server.MapPath("~/Images/maps/" + ImageName);

                    file.SaveAs(physicalPath); // save image in folder

                //var uri = new Uri(physicalPath, UriKind.Absolute); //ลบไฟล์รุป
                //System.IO.File.Delete(uri.LocalPath);

            }
            return RedirectToAction("index", "Student", new { user_id_title = ti.user_id_title, name_title = ti.name_title});
        }
        //------------------------------------------------------------------------------
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
    }
}