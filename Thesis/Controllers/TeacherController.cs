using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Cooperative_Rmutt.Models;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Web.Mvc;
using System.IO;
using System.Web;
using Project_Test1.Models;
using System.Net;
using System.Net.Mail;
using MySql.Data.MySqlClient;

namespace Cooperative_Rmutt.Controllers
{

    public class TeacherController : Controller
    {
        //string strConnection = "Data Source=JES-42COM\\SQLEXPRESS;Initial Catalog=Cooperative;Integrated Security=True;"; //Connect Database

        string strConnection = "Data Source=localhost;username=root;password=;database=coperativermutt;";//Connect Database

        public static string id;
        public ActionResult Index(TitileAll ti)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                ////////////////////////ดึงข้อมูลอาจารย์////////////////////////////////////////
                ///
                if (id == null)
                {
                    var sql_teacherdata = "Select * from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where teacher_id = '" + ti.user_id_title + "' ;";
                    conn.Open();
                    MySqlDataAdapter da_teacherdata = new MySqlDataAdapter(sql_teacherdata, conn);
                    DataSet ds_teacherdata = new DataSet();
                    da_teacherdata.Fill(ds_teacherdata, "TeachersandAuthorities");
                    DataTable dt_teacherdata = ds_teacherdata.Tables["TeachersandAuthorities"];
                    conn.Close();
                    id = ti.user_id_title;
                    @ViewData["name"] = dt_teacherdata.Rows[0]["name"].ToString();
                    @ViewData["department"] = dt_teacherdata.Rows[0]["major"].ToString();
                    @ViewData["teacher_id"] = dt_teacherdata.Rows[0]["teacher_id"].ToString();
                    @ViewData["Status"] = dt_teacherdata.Rows[0]["status"].ToString();
                    @ViewData["email"] = dt_teacherdata.Rows[0]["email"].ToString();
                    @ViewData["phone"] = dt_teacherdata.Rows[0]["phone"].ToString();
                }
                else
                {
                    var sql_teacherdata = "Select * from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where teacher_id = '" + id+ "' ;";
                    conn.Open();
                    MySqlDataAdapter da_teacherdata = new MySqlDataAdapter(sql_teacherdata, conn);
                    DataSet ds_teacherdata = new DataSet();
                    da_teacherdata.Fill(ds_teacherdata, "TeachersandAuthorities");
                    DataTable dt_teacherdata = ds_teacherdata.Tables["TeachersandAuthorities"];
                    conn.Close();
                    @ViewData["name"] = dt_teacherdata.Rows[0]["name"].ToString();
                    @ViewData["department"] = dt_teacherdata.Rows[0]["major"].ToString();
                    @ViewData["teacher_id"] = dt_teacherdata.Rows[0]["teacher_id"].ToString();
                    @ViewData["Status"] = dt_teacherdata.Rows[0]["status"].ToString();
                    @ViewData["email"] = dt_teacherdata.Rows[0]["email"].ToString();
                    @ViewData["phone"] = dt_teacherdata.Rows[0]["phone"].ToString();
                }
                /////////////////////////เช็คแบบสอบถาม/////////////////////////////////////
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ;";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                DataSet ds_displaycheckark = new DataSet();
                conn.Close();

                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                        string sql_setdate = "select *from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on t.teacher_id = a.user_id where a.major = '"+ @ViewData["department"] + "' and r.semester_id = '"+SEMESTER_ID+"'; select * from RequirementofTeachers Where semester_id = '" + SEMESTER_ID + "' and teacher_id = '" + ViewData["teacher_id"] + "'; select m.date_supervision,c.name_th,c.name_en,s.name_th,s.student_id from `match` m inner join Company c on m.company_id = c.company_id inner join Student s on m.student_id = s.student_id where teacher_id = '" + ViewData["teacher_id"] + "' and m.semester_id = '"+SEMESTER_ID+"'; select student_id,score16 from Document;";
                        conn.Open();
                        MySqlDataAdapter da_setdate = new MySqlDataAdapter(sql_setdate, conn);
                        DataSet ds_setdate = new DataSet();
                        da_setdate.Fill(ds_setdate, "RequirementSetDate");
                        DataTable dt_setdate = ds_setdate.Tables[0];
                        DataTable dt_require = ds_setdate.Tables[1];
                        conn.Close();
                        if (dt_setdate.Rows.Count > 0)
                        {
                            DateTime startsetdate = Convert.ToDateTime(dt_setdate.Rows[0]["startdate"]);
                            DateTime endsetdate = Convert.ToDateTime(dt_setdate.Rows[0]["enddate"]);
                            TimeSpan StaysetDay1 = startsetdate - nowdate;
                            TimeSpan StaysetDay2 = endsetdate - nowdate;

                            int DaysetStay1 = Convert.ToInt16(StaysetDay1.Days);
                            int DaysetStay2 = Convert.ToInt16(StaysetDay2.Days);

                            if (DaysetStay1 <= 0 && DaysetStay2 >= 0)
                            {
                                if (dt_require.Rows.Count > 0)
                                {
                                    string sql_updatetime = @"UPDATE TeachersandAuthorities  SET status = '" + "พร้อมออกนิเทศ" + "' Where teacher_id ='" + @ViewData["teacher_id"] + "';";
                                    conn.Open();
                                    MySqlDataAdapter da_updatetime = new MySqlDataAdapter(sql_updatetime, conn);
                                    DataSet ds_updatetime = new DataSet();
                                    da_updatetime.Fill(ds_updatetime, "`match`");
                                    conn.Close();
                                    ViewData["setdate"] = "Readyonly";
                                }
                                else
                                {
                                    ViewData["setdate"] = "Ontime";
                                }
                                return View(ds_setdate);

                            }
                            else
                            {
                                if (dt_require.Rows.Count > 0)
                                {
                                    string sql_updatetime = @"UPDATE TeachersandAuthorities  SET status = '" + "พร้อมออกนิเทศ" + "' Where teacher_id ='" + @ViewData["teacher_id"] + "';";
                                    conn.Open();
                                    MySqlDataAdapter da_updatetime = new MySqlDataAdapter(sql_updatetime, conn);
                                    DataSet ds_updatetime = new DataSet();
                                    da_updatetime.Fill(ds_updatetime, "`match`");
                                    conn.Close();
                                    
                                    ViewData["setdate"] = "Readyonly";
                                }
                                else
                                {
                                    string sql_updatetime = @"UPDATE TeachersandAuthorities  SET status = '" + "ไม่ประสงค์ออกนิเทศ" + "' Where teacher_id ='" + @ViewData["teacher_id"] + "';";
                                    conn.Open();
                                    MySqlDataAdapter da_updatetime = new MySqlDataAdapter(sql_updatetime, conn);
                                    DataSet ds_updatetime = new DataSet();
                                    da_updatetime.Fill(ds_updatetime, "`match`");
                                    conn.Close();
                                    @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
                                    ViewData["setdate"] = "NotReadyonly";
                                }
                                return View(ds_setdate);
                            }
                        }
                        else
                        {
                            ViewData["setdate"] = "NoSet";
                            return View(ds_setdate);
                        }

                    }

                }
                ViewData["setdate"] = "NoSet";
                ViewData["semester_null"] = "No";
                return View();
            }
        }

        public ActionResult Ask(string Ta_id, string department, string status,string name,string email,string phone)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ;" ;
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                conn.Close();
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["department"] = department;
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                @ViewData["phone"] = phone;
                if (status == "ไม่ประสงค์ออกนิเทศ")
                {
                    @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
                }
                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                        var sql_setdate = "select *from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on t.teacher_id = a.user_id where a.major = '" + @ViewData["department"] + "' and r.semester_id = '" + SEMESTER_ID + "';";
                        conn.Open();
                        MySqlDataAdapter da_setdate = new MySqlDataAdapter(sql_setdate, conn);
                        DataSet ds_setdate = new DataSet();
                        da_setdate.Fill(ds_setdate, "RequirementSetDate");
                        DataTable dt_setdate = ds_setdate.Tables[0];
                        conn.Close();
                        if (dt_setdate.Rows.Count > 0)
                        {
                            DateTime startsetdate = Convert.ToDateTime(dt_setdate.Rows[0]["startdate"]);
                            DateTime endsetdate = Convert.ToDateTime(dt_setdate.Rows[0]["enddate"]);
                            TimeSpan StaysetDay1 = startsetdate - nowdate;
                            TimeSpan StaysetDay2 = endsetdate - nowdate;

                            int DaysetStay1 = Convert.ToInt16(StaysetDay1.Days);
                            int DaysetStay2 = Convert.ToInt16(StaysetDay2.Days);

                            if (DaysetStay1 <= 0 && DaysetStay2 >= 0)
                            {
                                /////////////////////////////////แสดงข้อมูลแบบสอบถาม/////////////////////////////////////
                                SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                                var sql_displayask = "select c.company_id, c.name_th , s.name_th,s.student_id from Student s inner join Company c on s.company_id = c.company_id  inner join Account a on s.student_id = a.user_id where a.major = '" + department + "' and s.semester_out = '" + SEMESTER_ID + "' and s.confirm_status = 'รับแล้ว';  select distinct c.company_id, c.name_th,c.name_en from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id where s.semester_out = '" + SEMESTER_ID + "' and a.major = '" + department + "'and s.confirm_status = 'รับแล้ว'; select * from RequirementofTeachers Where semester_id = '" + SEMESTER_ID + "' and teacher_id = '" + Ta_id + "';";
                                conn.Open();
                                MySqlDataAdapter da_displayask = new MySqlDataAdapter(sql_displayask, conn);
                                DataSet ds_displayask = new DataSet();
                                da_displayask.Fill(ds_displayask, "Student");
                                DataTable dt_checkretime = ds_displayask.Tables[2];
                                conn.Close();
                                if (dt_checkretime.Rows.Count > 0)
                                {
                                    ViewData["date_false"] = "No";
                                    return View();
                                }
                                return View(ds_displayask);
                            }
                            else
                            {
                                SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                                var sql_checkretime = "select * from RequirementofTeachers Where semester_id = '" + SEMESTER_ID + "' and teacher_id = '" + Ta_id + "';";
                                conn.Open();
                                MySqlDataAdapter da_checkretime = new MySqlDataAdapter(sql_checkretime, conn);
                                DataSet ds_checkretime = new DataSet();
                                da_checkretime.Fill(ds_checkretime, "RequirementofTeachers");
                                DataTable dt_checkretime = ds_checkretime.Tables[0];
                                conn.Close();
                                if (dt_checkretime.Rows.Count > 0)
                                {
                                    ViewData["date_false"] = "No";
                                    return View();
                                }
                                else
                                {

                                    ViewData["date_false"] = "Yes";
                                    return View();
                                }

                            }
                        }
                        else
                        {
                            ViewData["date"] = "No";
                            return View();
                        }
                    }
                }
                ViewData["semester_null"] = "No";
                return View();
            }
        }
        [HttpPost]
        public ActionResult Ask(string Ta_id, TeacherRequirementModel ta)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ;";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                conn.Close();
                for (int j = 0; j < dt_semester.Rows.Count; j++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[j]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[j]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[j]["semester_id"].ToString();
                        int count_company = ta.Level.Length;
                        int count_day = ta.Day.Length;
                        string[] company_id = new string[count_company];
                        string[] level = new string[count_company];
                        string day = "";

                        ///////////////ใส่ข้อมูลวันที่เลือกลงไว้ใน Day ///////////////

                        for (int i = 0; i < count_day; i++)
                        {
                            if (day == "" || day == null)
                            {
                                day = ta.Day[i];
                            }
                            else
                            {
                                day = day + "," + ta.Day[i];
                            }

                        }
                        ///////////////วนตาม level ที่ใส่เข้ามา//////////////////
                        for (int i = 0; i < count_company; i++)
                        {
                            string sql_checkrow = "select * from RequirementofTeachers;"; /////////ดึงข้อมูลเพื่อเช็คแถวตลอด///////////
                            conn.Open();
                            MySqlDataAdapter da_checkrow = new MySqlDataAdapter(sql_checkrow, conn);
                            DataSet ds_checkrow = new DataSet();
                            da_checkrow.Fill(ds_checkrow, "RequirementofTeachers");
                            conn.Close();
                            DataTable re = ds_checkrow.Tables[0];
                            string[] split_level_company = ta.Level[i].Split('-');
                            company_id[i] = split_level_company[1];
                            level[i] = split_level_company[0];
                            if (level[i] == "0")
                            {

                            }
                            else
                            {
                                ////////////////////บันทึกข้อมูลลงในตารางแบบสอบถาม//////////////////
                                string sql_insertre = @"INSERT INTO RequirementofTeachers (requirement_id, teacher_id, company_id, number, day, travel,semester_id) VALUES('" + (re.Rows.Count + 1) + "','" + Ta_id + "','" + company_id[i] + "','" + level[i] + "','" + day + "','" + ta.Travel + "','" + SEMESTER_ID + "');";
                                conn.Open();
                                MySqlDataAdapter da_insertre = new MySqlDataAdapter(sql_insertre, conn);
                                DataSet ds_insertre = new DataSet();
                                da_insertre.Fill(ds_insertre, "TeachersandAuthorities");
                                conn.Close();

                            }
                        }
                    }
                }
                return RedirectToAction("Index");
            }

        }
        public ActionResult EditPersonal(string Ta_id, string Name, string department,string email,string phone,string status)
        {
            if (status == "ไม่ประสงค์ออกนิเทศ")
            {
                @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
            }
            @ViewData["name"] = Name;
            @ViewData["department"] = department;
            @ViewData["teacher_id"] = Ta_id;
            @ViewData["email"] = email;
            ViewData["phone"] = phone;
            return View();
        }
        [HttpPost]
        public ActionResult EditPersonal(string Ta_id, TeacherEditPersonalModel ta,string department,string status)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
               

                string sql_editpersonal = @"UPDATE TeachersandAuthorities  SET name = '" + ta.Name + "' ,phone = '" +ta.Phone+ "', email = '"+ta.Email+"'   Where teacher_id ='" + Ta_id + "';";
                conn.Open();
                MySqlDataAdapter da_editpersonal = new MySqlDataAdapter(sql_editpersonal, conn);
                DataSet ds_editpersonal = new DataSet();
                da_editpersonal.Fill(ds_editpersonal, "TeachersandAuthorities");
                conn.Close();
            }
            
            return RedirectToAction("EditPersonal",new { Ta_id = Ta_id, department = department,Name = ta.Name,email = ta.Email,phone = ta.Phone, status = status });
        }
        public ActionResult Viewmatch(string Ta_id, string department,string status,string name,string email,string phone)
        {
            if (status == "ไม่ประสงค์ออกนิเทศ")
            {
                @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
            }
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                DataSet ds_displayask = new DataSet();
                conn.Close();
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["department"] = department;
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                @ViewData["phone"] = phone;
                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                        string sql_viewmatch = "select distinct c.company_id, c.name_th , m.date_supervision,c.name_en from `match` m inner join Company c on m.company_id = c.company_id where teacher_id = '" + Ta_id + "' and semester_id = '" + SEMESTER_ID + "'; select c.company_id, s.name_th,s.student_id from Student s inner join Company c on s.company_id = c.company_id  inner join Account a on s.student_id = a.user_id where a.major = '" + department + "' and s.semester_out = '" + SEMESTER_ID + "' and s.confirm_status ='รับแล้ว';";
                        conn.Open();
                        MySqlDataAdapter da_viewmatch = new MySqlDataAdapter(sql_viewmatch, conn);
                        DataSet ds_viewmatch = new DataSet();
                        da_viewmatch.Fill(ds_viewmatch, "match");
                        conn.Close();
                        @ViewData["teacher_id"] = Ta_id;
                        @ViewData["department"] = department;

                        return View(ds_viewmatch);
                    }
                }
                ViewData["semester_null"] = "No";
                return View();
            }
        }
        [HttpPost]
        public ActionResult Viewmatch(string Ta_id, string department, UpdatetimeModel update, TeacherRequirementModel ta, string name, string email, string phone, string status)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ;";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                DataSet ds_displayask = new DataSet();
                conn.Close();
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["department"] = department;
                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        int count_time = update.Time.Length;
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();


                        conn.Close();
                        for (int j = 0; j < count_time; j++)
                        {
                            var sql_checkdateequal = "select  s.email,s.name_th,m.date_supervision,t.name from `match` m inner join Student s on m.student_id = s.student_id inner join TeachersandAuthorities t on m.teacher_id = t.teacher_id  Where m.teacher_id = '" + Ta_id + "' and m.semester_id = '" + SEMESTER_ID + "' and m.company_id = '" + ta.Level[j] + "';";
                            conn.Open();
                            MySqlDataAdapter da_checkdateequal = new MySqlDataAdapter(sql_checkdateequal, conn);
                            DataSet ds_checkdateequal = new DataSet();
                            da_checkdateequal.Fill(ds_checkdateequal, "`match`");
                            conn.Close();
                            if (update.Time[j] == ds_checkdateequal.Tables[0].Rows[0]["date_supervision"].ToString())
                            {
                            }
                            else
                            {
                                da_checkdateequal.Fill(ds_checkdateequal, "`match`");
                                string sql_updatetime = @"UPDATE `match`  SET date_supervision = '" + update.Time[j] + "' Where teacher_id ='" + Ta_id + "' and semester_id = '" + SEMESTER_ID + "' and company_id = '" + ta.Level[j] + "';";
                                conn.Open();
                                MySqlDataAdapter da_updatetime = new MySqlDataAdapter(sql_updatetime, conn);
                                DataSet ds_updatetime = new DataSet();
                                da_updatetime.Fill(ds_updatetime, "`match`");
                                conn.Close();
                                var sql_sendmail = "select  s.email,s.name_th,m.date_supervision,t.name from `match` m inner join Student s on m.student_id = s.student_id inner join TeachersandAuthorities t on m.teacher_id = t.teacher_id  Where m.teacher_id = '" + Ta_id + "' and m.semester_id = '" + SEMESTER_ID + "' and m.company_id = '" + ta.Level[j] + "';";
                                conn.Open();
                                MySqlDataAdapter da_sendmail = new MySqlDataAdapter(sql_sendmail, conn);
                                DataSet ds_sendmail = new DataSet();
                                da_sendmail.Fill(ds_sendmail, "`match`");
                                conn.Close();
                                for (int k = 0; k < ds_sendmail.Tables[0].Rows.Count; k++)
                                {
                                    MailMessage mm = new MailMessage("EnCooperativeRmuttK6@gmail.com", ds_sendmail.Tables[0].Rows[k]["email"].ToString());
                                    mm.Subject = "แจ้งเตือนการเปลี่ยนวันนิเทศจาก " + name;
                                    mm.Body = "มีความประสงค์เปลี่ยนวันนิเทศเป็น " + ds_sendmail.Tables[0].Rows[k]["date_supervision"].ToString();
                                    mm.IsBodyHtml = false;

                                    SmtpClient smtp = new SmtpClient();
                                    smtp.Host = "smtp.gmail.com";
                                    smtp.Port = 587;
                                    smtp.EnableSsl = true;

                                    NetworkCredential nc = new NetworkCredential("EnCooperativeRmuttK6@gmail.com", "encoopk6");
                                    smtp.UseDefaultCredentials = false;
                                    smtp.Credentials = nc;
                                    smtp.Send(mm);
                                }

                            }
                        }


                    }
                }
            }
            return RedirectToAction("Viewmatch", new { Ta_id = ViewData["teacher_id"], Department = ViewData["department"], name = name, email = email, phone = phone, status = status });
        }
        public ActionResult FacData(string Ta_id, string department, string FacData, string status,string name,string email,string phone)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (status == "ไม่ประสงค์ออกนิเทศ")
                {
                    @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
                } 
                @ViewData["phone"] = phone;
                @ViewData["department"] = department;
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                if (FacData != null)
                {
                    string SEMESTER_ID;
                    var sql_semester = "select * from DefineSemester ;";
                    conn.Open();
                    MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                    DataSet ds_semester = new DataSet();
                    da_semester.Fill(ds_semester, "DefineSemester");
                    DataTable dt_semester = ds_semester.Tables[0];
                    DataSet ds_displaycheckark = new DataSet();
                    conn.Close();
                    for (int i = 0; i < dt_semester.Rows.Count; i++)
                    {
                        DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                        DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                        DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                        TimeSpan StayDay1 = startdate - nowdate;
                        TimeSpan StayDay2 = enddate - nowdate;

                        int DayStay1 = Convert.ToInt16(StayDay1.Days);
                        int DayStay2 = Convert.ToInt16(StayDay2.Days);

                        if (DayStay1 <= 0 && DayStay2 >= 0)
                        {
                            SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                            string sql_FacData = "select c.company_id, c.name_th , s.name_th,s.mobile_phone,s.student_id from Student s inner join Company c on s.company_id = c.company_id  inner join Account a on s.student_id = a.user_id where a.major = '" + department + "' and s.semester_out = '" + SEMESTER_ID + "' and c.company_id = '" + FacData + "' and s.confirm_status ='รับแล้ว';  select distinct  c.company_id,c.name_th,c.name_en,c.business_type,c.job_day,c.job_time,c.phone,c.fax,c.website,c.email,ad.number,ad.building,ad.floor,ad.lane,ad.road,t.sub_area,t.area,t.province,t.postal_code,ad.px,ad.py from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on c.address_id = ad.address_id inner join Prototypeaddress t on ad.prototype_id = t.prototype_id where semester_out = '" + SEMESTER_ID + "' and a.major = '" + department + "' and c.company_id = '" + FacData + "'and s.confirm_status ='รับแล้ว';";
                            conn.Open();
                            MySqlDataAdapter da_FacData = new MySqlDataAdapter(sql_FacData, conn);
                            DataSet ds_FacData = new DataSet();
                            da_FacData.Fill(ds_FacData, "Student");
                            conn.Close();
                           
                            return View(ds_FacData);
                        }
                    }
                    ViewData["semester_null"] = "No";
                }
                else
                {
                    string SEMESTER_ID;
                    var sql_semester = "select * from DefineSemester ;";
                    conn.Open();
                    MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                    DataSet ds_semester = new DataSet();
                    da_semester.Fill(ds_semester, "DefineSemester");
                    DataTable dt_semester = ds_semester.Tables[0];
                    DataSet ds_displaycheckark = new DataSet();
                    conn.Close();
                    for (int i = 0; i < dt_semester.Rows.Count; i++)
                    {
                        DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                        DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                        DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                        TimeSpan StayDay1 = startdate - nowdate;
                        TimeSpan StayDay2 = enddate - nowdate;

                        int DayStay1 = Convert.ToInt16(StayDay1.Days);
                        int DayStay2 = Convert.ToInt16(StayDay2.Days);

                        if (DayStay1 <= 0 && DayStay2 >= 0)
                        {
                            SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                            string sql_FacData = "select c.company_id, c.name_th , s.name_th,s.mobile_phone,student_id from Student s inner join Company c on s.company_id = c.company_id  inner join Account a on s.student_id = a.user_id where a.major = '" + department + "' and s.semester_out = '" + SEMESTER_ID + "' and s.confirm_status ='รับแล้ว';  select distinct  c.company_id,c.name_th,c.name_en,c.business_type,c.job_day,c.job_time,c.phone,c.fax,c.website,c.email,ad.number,ad.building,ad.floor,ad.lane,ad.road,t.sub_area,t.area,t.province,t.postal_code,ad.px,ad.py from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on c.address_id = ad.address_id inner join Prototypeaddress t on ad.prototype_id = t.prototype_id where semester_out = '" + SEMESTER_ID + "' and a.major = '" + department + "' and s.confirm_status ='รับแล้ว';";
                            conn.Open();
                            MySqlDataAdapter da_FacData = new MySqlDataAdapter(sql_FacData, conn);
                            DataSet ds_FacData = new DataSet();
                            da_FacData.Fill(ds_FacData, "Student");
                            conn.Close();
                            
                            return View(ds_FacData);
                        }
                    }
                    ViewData["semester_null"] = "No";
                }
            }
            return View();
        }
        [HttpPost]
        public ActionResult Status(string status, string Ta_id)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string sql_updatetime = @"UPDATE TeachersandAuthorities  SET status = '" + status + "' Where teacher_id ='" + Ta_id + "';";
                conn.Open();
                MySqlDataAdapter da_updatetime = new MySqlDataAdapter(sql_updatetime, conn);
                DataSet ds_updatetime = new DataSet();
                da_updatetime.Fill(ds_updatetime, "`match`");
                conn.Close();
                return RedirectToAction("Index");
            }
        }
        public ActionResult Score(string Ta_id, string department, string status,string name,string email,string phone)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ;";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                DataSet ds_displayask = new DataSet();
                conn.Close();
                if (status == "ไม่ประสงค์ออกนิเทศ")
                {
                    @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
                }
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["department"] = department;
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                @ViewData["phone"] = phone;
                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                        string sql_viewmatch = "select c.name_th,c.name_en,s.name_th,s.student_id,c.company_id from `match` m inner join Company c on m.company_id = c.company_id inner join Student s on m.student_id = s.student_id where teacher_id = '" + ViewData["teacher_id"] + "'and m.semester_id = '" + SEMESTER_ID + "'; select student_id,score16 from Document";
                        conn.Open();
                        MySqlDataAdapter da_viewmatch = new MySqlDataAdapter(sql_viewmatch, conn);
                        DataSet ds_viewmatch = new DataSet();
                        da_viewmatch.Fill(ds_viewmatch, "match");
                        conn.Close();
                        

                        return View(ds_viewmatch);
                    }
                }
                ViewData["semester_null"] = "No";
                return View();
            }

        }
        public ActionResult Rate(string student,string Ta_id, string department,string status,string name,string email,string phone)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string sql_Rate = " select  s.name_th,a.major,c.name_th from Student s inner join Account a on a.user_id = s.student_id inner join Company c on c.company_id = s.company_id where student_id = '" + student + "';select score16 from Document where student_id ='" + student + "';select title_th,title_en,score16 from Document where student_id ='" + student + "';";
                conn.Open();
                MySqlDataAdapter da_Rate = new MySqlDataAdapter(sql_Rate, conn);
                DataSet ds_Rate = new DataSet();
                da_Rate.Fill(ds_Rate, "Student");
                conn.Close();
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                @ViewData["phone"] = phone;
                if (ds_Rate.Tables[1].Rows.Count > 0 && ds_Rate.Tables[1].Rows[0]["score16"].ToString() != "")
                {
                    ViewData["student_id"] = "Do";
                    @ViewData["teacher_id"] = Ta_id;
                    @ViewData["department"] = department;
                    return View();
                }
                    DataTable dt_Rate = ds_Rate.Tables["Student"];
                    ViewData["student_name"] = dt_Rate.Rows[0]["name_th"];
                    ViewData["student_department"] = dt_Rate.Rows[0]["major"];
                    ViewData["company_name"] = dt_Rate.Rows[0][2];
                    if (ds_Rate.Tables[2].Rows.Count > 0 && ds_Rate.Tables[2].Rows[0]["title_th"].ToString() != "")
                    {
                    ViewData["Title_th"] = ds_Rate.Tables[2].Rows[0]["title_th"];
                    ViewData["Title_en"] = ds_Rate.Tables[2].Rows[0]["title_en"];
                    }
                    else
                    {
                    ViewData["Title_th"] = "null";
                    }
                    ViewData["student_id"] = student;
                    @ViewData["teacher_id"] = Ta_id;
                    @ViewData["department"] = department;
                return View();
            }
        }
        [HttpPost]
        public ActionResult Rate(string student,string More, TeacherRateStudent sc)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (More == "")
                {
                    //string sql_insertrate = @"UPDATE  (student_id,score16) VALUES('" + student + "','" + sc.score_1+"," + sc.score_2 + ","+ sc.score_3 + ","+ sc.score_4 + ","+ sc.score_5 + ","+ sc.score_6 + ","+ sc.score_7 + ","+ sc.score_8 + ","+ sc.score_9 + ","+ sc.score_10 + "')";
                    string sql_insertrate = @"UPDATE  Document  SET score16 = '" + sc.score_1 + "," + sc.score_2 + "," + sc.score_3 + "," + sc.score_4 + "," + sc.score_5 + "," + sc.score_6 + "," + sc.score_7 + "," + sc.score_8 + "," + sc.score_9 + "," + sc.score_10 + "'  Where student_id ='" + student + "';";
                    conn.Open();
                    MySqlDataAdapter da_insertrate = new MySqlDataAdapter(sql_insertrate, conn);
                    DataSet ds_insertrate = new DataSet();
                    da_insertrate.Fill(ds_insertrate, "TeachersandAuthorities");
                    conn.Close();
                }
                else
                {
                    //string sql_insertrate = @"UPDATE INTO Document (student_id,score16,comment) VALUES('" + student + "','" + sc.score_1 + "," + sc.score_2 + "," + sc.score_3 + "," + sc.score_4 + "," + sc.score_5 + "," + sc.score_6 + "," + sc.score_7 + "," + sc.score_8 + "," + sc.score_9 + "," + sc.score_10 + "','" + More + "')";
                    //string sql_insertrate = @"UPDATE INTO Document (student_id,score16,comment) VALUES('" + student + "','" + sc.score_1 + "," + sc.score_2 + "," + sc.score_3 + "," + sc.score_4 + "," + sc.score_5 + "," + sc.score_6 + "," + sc.score_7 + "," + sc.score_8 + "," + sc.score_9 + "," + sc.score_10 + "','" + More + "')";
                    string sql_insertrate = @"UPDATE  Document  SET score16 = '" + sc.score_1 + "," + sc.score_2 + "," + sc.score_3 + "," + sc.score_4 + "," + sc.score_5 + "," + sc.score_6 + "," + sc.score_7 + "," + sc.score_8 + "," + sc.score_9 + "," + sc.score_10 + "'  Where student_id ='" + student + "'; UPDATE Document SET comment ='" + More + "'Where student_id ='" + student + "';";
                    conn.Open();
                    MySqlDataAdapter da_insertrate = new MySqlDataAdapter(sql_insertrate, conn);
                    DataSet ds_insertrate = new DataSet();
                    da_insertrate.Fill(ds_insertrate, "TeachersandAuthorities");
                    conn.Close();
                }

            }
            return RedirectToAction("Index");
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

        public ActionResult sk13(string company_id, string department,string Ta_id)
        {
            string fileName = "20160702-12.pdf";
            string sourcePath = Server.MapPath("~/pdf_file");
            string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, Ta_id + "-" + fileName);
            System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

            string oldFile = Server.MapPath("~/pdf_file/" + fileName);
            string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + Ta_id + "-" + fileName);

            //////////////////////////////////////////////////////////////////////////////
            string[] Student;
            string[] Major;
            
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string SEMESTER_ID;
                var sql_semester = "select * from DefineSemester ";
                conn.Open();
                MySqlDataAdapter da_semester = new MySqlDataAdapter(sql_semester, conn);
                DataSet ds_semester = new DataSet();
                da_semester.Fill(ds_semester, "DefineSemester");
                DataTable dt_semester = ds_semester.Tables[0];
                DataSet ds_displayask = new DataSet();
                conn.Close();

                for (int i = 0; i < dt_semester.Rows.Count; i++)
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt_semester.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt_semester.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        SEMESTER_ID = dt_semester.Rows[i]["semester_id"].ToString();
                        string sql_viewmatch = "select c.company_id, s.name_th,a.major,s.gender from Student s inner join Company c on s.company_id = c.company_id  inner join Account a on s.student_id = a.user_id where a.major = '" + department + "' and s.semester_out = '" + SEMESTER_ID + "' and s.company_id = '" + company_id + "'and s.confirm_status ='รับแล้ว'; select c.name_th,c.name_en,a.number,a.building,a.floor,a.lane,a.road,t.sub_area,t.area,t.province,t.postal_code,c.phone,c.fax,c.email from Company c inner join Address a on c.address_id = a.address_id inner join Prototypeaddress t on a.prototype_id = t.prototype_id  where c.company_id = '" + company_id+ "';select t.name from TeachersandAuthorities t inner join `match` m on t.teacher_id =  m.teacher_id where m.company_id = '" + company_id + "' and m.teacher_id = '" + Ta_id + "';";
                        conn.Open();
                        MySqlDataAdapter da_viewmatch = new MySqlDataAdapter(sql_viewmatch, conn);
                        DataSet ds_viewmatch = new DataSet();
                        da_viewmatch.Fill(ds_viewmatch, "Student");
                        conn.Close();
                        Student = new String[ds_viewmatch.Tables[0].Rows.Count];
                        Major = new String[ds_viewmatch.Tables[0].Rows.Count];
                        
                        for (int j = 0; j < ds_viewmatch.Tables[0].Rows.Count; j++)
                        {
                            Student[j] = ds_viewmatch.Tables[0].Rows[j]["name_th"].ToString();
                            Major[j] = ds_viewmatch.Tables[0].Rows[j]["major"].ToString();
                           
                        }
                        using (var reader = new PdfReader(oldFile))
                        {
                            using (var fileStream = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                            {
                                var document = new Document(reader.GetPageSizeWithRotation(1));
                                var writer = PdfWriter.GetInstance(document, fileStream);

                                document.Open();

                                for (var k = 1; k <= Student.Length; k++) // ใส่ข้อมูล 2 หน้า
                                {
                                    // ใส่ข้อมูลหน้า 1
                                    if (k == 1)
                                    {
                                        string text;
                                        document.NewPage();

                                        BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                        PdfImportedPage importedPage = writer.GetImportedPage(reader, k);

                                        PdfContentByte cb = writer.DirectContent;
                                        cb.SetColorFill(BaseColor.BLACK);
                                        cb.SetFontAndSize(baseFont, 12);

                                        cb.AddTemplate(importedPage, 0, 0);
                                        if (ds_viewmatch.Tables[1].Rows.Count > 0)
                                        {
                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["name_th"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 120, 618, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["name_en"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 130, 599, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["number"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 122, 581, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["building"].ToString(); ;
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 254, 581, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["floor"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 324, 581, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["lane"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 394, 581, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["road"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 500, 581, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["sub_area"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 143, 563, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["area"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 275, 563, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["province"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 395, 563, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["postal_code"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 524, 563, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["phone"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 137, 545, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["fax"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 267, 545, 0);
                                            cb.EndText();

                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[1].Rows[0]["email"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 355, 545, 0);
                                            cb.EndText();
                                        }
                                        if (ds_viewmatch.Tables[2].Rows.Count > 0)
                                        {
                                            cb.BeginText();
                                            text = ds_viewmatch.Tables[2].Rows[0]["name"].ToString();
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 435, 222, 0);
                                            cb.EndText();
                                        }
                                        /////////////////////////////////////////////////////////////////// รายชื่อนักศึกษา


                                        int line = 471;

                                        for (int m = 0; m < Student.Length; m++)
                                        {
                                            
                                                cb.BeginText();
                                                text =  Student[m];
                                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 90, line, 0);
                                                cb.EndText();
                                            
                                            

                                            cb.BeginText();
                                            text = Major[m];
                                            cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 339, line, 0);
                                            cb.EndText();

                                            line = line - 22;

                                        }

                                      

                                    }
                                    if (k != 0)
                                    {
                                       
                                            document.NewPage();

                                            PdfImportedPage importedPage = writer.GetImportedPage(reader, 2);

                                            PdfContentByte cb = writer.DirectContent;

                                            cb.AddTemplate(importedPage, 0, 0);
                                       
                                            document.NewPage();

                                            PdfImportedPage importedPage1 = writer.GetImportedPage(reader, 3);

                                            PdfContentByte cb1 = writer.DirectContent;

                                            cb1.AddTemplate(importedPage1, 0, 0);
                                        
                                    }
                                }

                                document.Close();
                                writer.Close();
                            }
                        }
                        return RedirectToAction("GetReport", "Teacher", new { doc_id = Ta_id + "-" + fileName });
                    }
                }

            }
            return View();
        }
        public ActionResult Studentdata(string Ta_id, string department, string status, string name, string email, string phone,string student)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (status == "ไม่ประสงค์ออกนิเทศ")
                {
                    @ViewData["Status"] = "ไม่ประสงค์ออกนิเทศ";
                }
                @ViewData["phone"] = phone;
                @ViewData["department"] = department;
                @ViewData["teacher_id"] = Ta_id;
                @ViewData["name"] = name;
                @ViewData["email"] = email;
                string sql_studentdata = "select * from Student where student_id ='"+student+ "';select * from Supervisor where student_id ='" + student + "';";
                conn.Open();
                MySqlDataAdapter da_studentdata = new MySqlDataAdapter(sql_studentdata, conn);
                DataSet ds_studentdata = new DataSet();
                da_studentdata.Fill(ds_studentdata, "Student");
                conn.Close();

                return View(ds_studentdata);   
                }  
           
        }
        public ActionResult Logout_teacher()
        {
            id = null;
            return RedirectToAction("LoginTemplate", "Home");
        }
    }
}
               