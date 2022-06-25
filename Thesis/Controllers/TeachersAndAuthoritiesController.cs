using iTextSharp.text;
using iTextSharp.text.pdf;
using LumenWorks.Framework.IO.Csv;
using Project_03.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Project_Test1.Models;
using MySql.Data.MySqlClient;
namespace Project_03.Controllers
{
    public class TeachersAndAuthoritiesController : Controller
    {
        //string strConnection = "Data Source=JES-42COM\\SQLEXPRESS;Initial Catalog=Cooperative;Integrated Security=True;"; //Connect Database

        //string strConnection = "Data Source=localhost;username=root;password=;database=coperativermutt;";//Connect Database

        public static string TA_ID;
        public static string TA_MAJOR;
        public static string TA_NAME;
        public static string TA_PROFILE;
        public static string TA_POSITION;
        public static string SEMESTER_ID;

        public static string SEM;

        // GET: TeacherAndAuthorities

        public ActionResult TeacherHome(string id)
        {           
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "SELECT t.teacher_id, t.name, a.position, a.major FROM TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where teacher_id = '" + id + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();
                
                TA_ID = dt.Rows[0][0].ToString();
                TA_MAJOR = dt.Rows[0][3].ToString();
                TA_NAME = dt.Rows[0][1].ToString();
                TA_PROFILE = TA_ID + ".jpg";
                TA_POSITION = dt.Rows[0][2].ToString();

                ViewData["teacher_id"] = TA_ID;
                ViewData["profile_id"] = TA_PROFILE; 
                ViewData["name"] = TA_NAME;
                ViewData["department"] = TA_MAJOR;

                SEMESTER_ID = getTerm(); // ตรวจสอบว่าเป็นเทอมไหน ปีไหน

                ViewData["semCheck"] = SEMESTER_ID;

                if (SEMESTER_ID == null)
                {
                    return View();
                }
                else
                {
                    //var sql2 = "select * from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "'";

                    //conn.Open();
                    //MySqlDataAdapter da2 = new MySqlDataAdapter(sql2, conn);
                    //DataSet ds2 = new DataSet();
                    //da2.Fill(ds2, "RequirementSetDate");
                    //DataTable dt2 = ds2.Tables[0];
                    //conn.Close();

                    //if (dt2.Rows.Count == 0)
                    //{
                    //    ViewData["set_date_requirement"] = "ยังไม่ได้เลือก";
                    //    return View();
                    //}
                    //else
                    //{
                    //    ViewData["set_date_requirement"] = dt2.Rows[0][4].ToString();
                    //    return View();
                    //}

                    string[] split_sem = SEMESTER_ID.Split('-');
                    string term = split_sem[1];
                    string year = split_sem[0];

                    ViewData["term-sem"] = "เทอม " + term + " ปีการศึกษา " + year;

                    var sql2 = "select r.enddate from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "';";
                    sql2 += "select t.teacher_id, t.name from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where a.major = '" + TA_MAJOR + "' and t.status = 'พร้อมออกนิเทศ' and a.position = 'อาจารย์นิเทศ';";
                    sql2 += "select distinct t.teacher_id, t.name from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '"+SEMESTER_ID+"';";
                    sql2 += "select s.student_id, s.name_th from `match` m inner join Student s on m.student_id = s.student_id inner join Account a on a.user_id = s.student_id where m.semester_id = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "';";
                    sql2 += "select s.student_id, s.name_th, c.name_th, t.name , m.semester_id from Document d inner join `match` m on d.student_id = m.student_id inner join Student s on s.student_id = m.student_id inner join Company c on c.company_id = s.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and m.semester_id = '" + SEMESTER_ID + "' order by t.name";

                    DataSet ds2 = getDatabase(sql2);

                    if (ds2.Tables[0].Rows.Count == 0) //// กำหนดวันส่งแบบสอบถาม
                    {
                        ViewData["set_date_requirement"] = "ยังไม่ได้เลือก";
                    }
                    else
                    {
                        ViewData["set_date_requirement"] = ds2.Tables[0].Rows[0][0].ToString();
                    }

                    if(ds2.Tables[1].Rows.Count > ds2.Tables[2].Rows.Count) //// ส่งแบบสอบถามครบยัง
                    {
                        ViewData["Requilement"] = "ยังไม่ครบ";
                    }
                    else
                    {
                        ViewData["Requilement"] = "เรียบร้อย";
                    }

                    if (ds2.Tables[3].Rows.Count == 0) //// จับคู่ยัง
                    {
                        ViewData["matching"] = "ยังไม่ทำการคัดเลือก";
                    }
                    else
                    {
                        ViewData["matching"] = "เรียบร้อย";
                    }

                    if (ds2.Tables[3].Rows.Count > ds2.Tables[4].Rows.Count) //// ส่งสก.16 ครบยัง
                    {
                        ViewData["sk16"] = "ยังไม่ครบ";
                    }
                    else
                    {
                        ViewData["sk16"] = "เรียบร้อย";
                    }

                    return View();

                }
            }
        }

        public ActionResult TeacherEdit()
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "SELECT a.user_id, a.password, t.email, t.phone FROM TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where teacher_id = '"+TA_ID+"';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables["TeachersandAuthorities"];
                conn.Close();

                TA_ID = dt.Rows[0][0].ToString();

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["pwd"] = dt.Rows[0][1].ToString();
                ViewData["email"] = dt.Rows[0][2].ToString();
                ViewData["phone"] = dt.Rows[0][3].ToString();
                ViewData["semCheck"] = SEMESTER_ID;

                return View();

            }
        }

        [HttpPost]
        public ActionResult TeacherEdit(Teachers_And_Authorities ta)
        {
            string UpdateDate = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (ta.pwd == null || ta.pwd == "")
                {
                    return View();
                }
                else
                {
                    string sql = @"UPDATE TeachersandAuthorities set name = '" + ta.name + @"',email = '" + ta.email + @"',phone = '" + ta.phone + @"' where teacher_id = '" + TA_ID + "'; update Account set password = '" + ta.pwd + @"', date_time_update = '" + UpdateDate + @"' where user_id = '" + TA_ID + "'; select name from TeachersandAuthorities where teacher_id = '" + TA_ID + "';";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");
                    DataTable dt = ds.Tables["TeachersandAuthorities"];
                    conn.Close();

                    TA_NAME = dt.Rows[0][0].ToString();

                    return RedirectToAction("TeacherEdit", "TeachersAndAuthorities", new { id = TA_ID});
                }
            }
        }

        public ActionResult Teachermatching()
        {
            string[] split_id = SEMESTER_ID.Split('-');
            string year = split_id[0];
            string term = split_id[1];

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                /// เปลี่ยนแล้ว //var sql = "select distinct c.company_id, c.name_th, t.teacher_id, t.name, m.date from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id where s.major = '" + TA_MAJOR + "' and t.department = '" + TA_MAJOR + "' and t.status = 'พร้อมออกนิเทศ' and s.semester_id = '" + SEMESTER_ID + "' and m.semester_id = '"+SEMESTER_ID+"'";
                //var sql = "select distinct c.company_id, c.name_th, t.teacher_id, t.name, m.date from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a on a.user_id = s.student_id inner join Account a1 on a1.user_id = t.teacher_id where a.major = '"+TA_MAJOR+"' and a1.major = '"+TA_MAJOR+"' and t.status = 'พร้อมออกนิเทศ' and s.semester_id = '"+SEMESTER_ID+"' and m.semester_id = '"+SEMESTER_ID+"';";
                var sql = "select distinct c.company_id, c.name_th, t.teacher_id, t.name, m.date from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a on a.user_id = s.student_id inner join Account a1 on a1.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and a1.major = '" + TA_MAJOR + "' and t.status = 'พร้อมออกนิเทศ' and s.semester_out = '" + SEMESTER_ID + "' and m.semester_id = '" + SEMESTER_ID + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable[] d = new DataTable[ds.Tables[0].Rows.Count];

                conn.Close();

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    /// เปลี่ยนแล้ว //var sql1 = "select student_id,company_id,name_th,semester_id from Student where company_id = '" + ds.Tables[0].Rows[i][0].ToString() + "' and semester_id = '" + SEMESTER_ID + "' and major = '" + TA_MAJOR + "'";
                    //var sql1 = "select s.student_id,s.company_id,s.name_th,s.semester_id from Student s inner join Account a on s.student_id = a.user_id where s.company_id = '" + ds.Tables[0].Rows[i][0].ToString() + "' and s.semester_id = '"+SEMESTER_ID+"' and a.major = '"+TA_MAJOR+"'; ";

                    //var sql1 = "select s.student_id,s.company_id,s.name_th,s.semester_id from Student s inner join Account a on s.student_id = a.user_id where s.company_id = '" + ds.Tables[0].Rows[i][0].ToString() + "' and s.semester_out = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "'; ";

                    var sql1 = "select s.student_id,s.company_id,s.name_th,s.semester_id from Student s inner join Account a on s.student_id = a.user_id inner join `match` m on m.student_id = s.student_id where s.company_id = '" + ds.Tables[0].Rows[i][0].ToString() + "' and s.semester_out = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "'; ";

                    conn.Open();
                    MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "www");
                    conn.Close();
                    DataRow reRow;
                    d[i] = new DataTable("DataTable " + i);
                    d[i].Columns.Add("Name", typeof(string));

                    for (int j = 0; j < ds1.Tables[0].Rows.Count; j++)
                    {
                        reRow = d[i].NewRow();
                        reRow[0] = ds1.Tables[0].Rows[j][2].ToString();
                        d[i].Rows.Add(reRow);
                       
                    }
                    ds.Tables.Add(d[i]);
                }

                //conn.Close();

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["term"] = "เทอม " + term + " ปีการศึกษา " + year;

                return View(ds);
            }
        }

        public ActionResult TeacherMatchingEdit(string company_id)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                //string[] split_id = TandC_id.Split('-');
                //string company_id = split_id[0];
                //string teacher_id = split_id[1];

                var sql = "select distinct m.company_id, c.name_th, m.teacher_id, t.name from `match` m inner join account a on m.teacher_id = a.user_id inner join TeachersandAuthorities t on m.teacher_id = t.teacher_id inner join Company c on c.company_id = m.company_id where m.semester_id = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "' and c.company_id = '" + company_id + "'; ";
                sql += "select m.company_id, c.name_th, m.student_id, s.name_th from `match` m inner join account a on m.student_id = a.user_id inner join Student s on s.student_id = a.user_id inner join Company c on c.company_id = m.company_id where m.semester_id = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "' and c.company_id = '" + company_id + "'; ";
                sql += "select distinct r.teacher_id , t.name from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on a.user_id = t.teacher_id where r.semester_id = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "'; ";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables["TeachersandAuthorities"];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["semCheck"] = SEMESTER_ID;

                return View(ds);
            }
        }

        [HttpPost]
        public ActionResult TeacherMatchingEditSubmit(string company, string teacher)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select m.match_id,m.company_id, c.name_th, m.teacher_id, t.teacher_id from `match` m inner join account a on m.teacher_id = a.user_id inner join TeachersandAuthorities t on t.teacher_id = a.user_id inner join Company c on c.company_id = m.company_id where m.semester_id = '" + SEMESTER_ID + "' and a.major = '" + TA_MAJOR + "' and c.company_id = '" + company + "'; ";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                conn.Close();

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    var ud = @"UPDATE `match` set teacher_id = '" + teacher + "', date = 'ยังไม่ได้กำหนด' , date_supervision = 'ยังไม่ได้กำหนด' where match_id = '" + ds.Tables[0].Rows[i][0].ToString() + "';";

                    conn.Open();
                    MySqlDataAdapter daUd = new MySqlDataAdapter(ud, conn);
                    DataSet dsUd = new DataSet();
                    daUd.Fill(dsUd, "TeachersandAuthorities");
                    conn.Close();
                }

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;

                return RedirectToAction("TeacherMatching", "TeachersAndAuthorities");
            }
        }

        public ActionResult erty()
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                //// แก้ละ /var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id where s.major = 'วิศวกรรมคอมพิวเตอร์'; select teacher_id, name from TeachersandAuthorities where department = 'วิศวกรรมคอมพิวเตอร์' and status = 'พร้อมออกนิเทศ' and position = 'อาจารย์นิเทศ'";
                //sql += "select r.requirement_id, t.teacher_id, t.name, c.company_id, c.name_th, r.number, r.day, r.travel from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id where t.department = 'วิศวกรรมคอมพิวเตอร์' and semester_id = '" + SEMESTER_ID + "'; select distinct t.teacher_id, t.name from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id where t.department = 'วิศวกรรมคอมพิวเตอร์' and semester_id = '" + SEMESTER_ID + "'; select distinct c.company_id, c.name_th from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id where t.department = 'วิศวกรรมคอมพิวเตอร์' and semester_id = '" + SEMESTER_ID + "';";

                //var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where a.major = '"+TA_MAJOR+"';";
                //    sql += "select t.teacher_id, t.name from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where a.major = '"+TA_MAJOR+"' and t.status = 'พร้อมออกนิเทศ' and t.position = 'อาจารย์นิเทศ';";
                //    sql += "select r.requirement_id, t.teacher_id, t.name, c.company_id, c.name_th, r.number, r.day, r.travel from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '"+TA_MAJOR+"' and r.semester_id = '"+SEMESTER_ID+"'";
                //    sql += "select distinct t.teacher_id, t.name from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '"+TA_MAJOR+"' and r.semester_id = '"+SEMESTER_ID+"'; ";
                //    sql += "select distinct c.company_id, c.name_th from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '"+TA_MAJOR+"' and r.semester_id = '"+SEMESTER_ID+"'; ";

                var sql = "select distinct c.company_id, c.name_th from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and s.confirm_status = 'รับแล้ว' and s.semester_out = '"+SEMESTER_ID+"';";
                sql += "select t.teacher_id, t.name from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where a.major = '" + TA_MAJOR + "' and t.status = 'พร้อมออกนิเทศ' and a.position = 'อาจารย์นิเทศ';";
                sql += "select r.requirement_id, t.teacher_id, t.name, c.company_id, c.name_th, r.number, r.day, r.travel from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "';";
                sql += "select distinct t.teacher_id, t.name from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "'; ";
                sql += "select distinct c.company_id, c.name_th from RequirementofTeachers r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Company c on r.company_id = c.company_id inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "'; ";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                conn.Close();

                DataTable dt_count_teacher = ds.Tables[1];
                DataTable dt = ds.Tables[2];
                DataTable dt_count_re_teacher = ds.Tables[3];
                //DataTable dt2 = ds.Tables[4];

                if (dt_count_teacher.Rows.Count >= dt_count_re_teacher.Rows.Count)
                {
                    DataTable req = new DataTable();
                    req.Columns.Add("teacher", typeof(string));
                    req.Columns.Add("level1", typeof(string));
                    req.Columns.Add("level2", typeof(string));
                    req.Columns.Add("day", typeof(string));
                    req.Columns.Add("travel", typeof(string));
                    DataRow reqRow;

                    string[] teacher_name = new string[dt_count_re_teacher.Rows.Count];
                    string[] teacher_id = new string[dt_count_re_teacher.Rows.Count];
                    string[] number1 = new string[dt_count_re_teacher.Rows.Count];
                    string[] number2 = new string[dt_count_re_teacher.Rows.Count];
                    string[] day = new string[dt_count_re_teacher.Rows.Count];
                    string[] travel = new string[dt_count_re_teacher.Rows.Count];

                    for (int i = 0; i < dt_count_re_teacher.Rows.Count; i++)
                    {
                        teacher_name[i] = dt_count_re_teacher.Rows[i][1].ToString();
                        teacher_id[i] = dt_count_re_teacher.Rows[i][0].ToString();

                        for (int j = 0; j < dt.Rows.Count; j++)
                        {
                            if (dt.Rows[j][1].ToString() == teacher_id[i] && dt.Rows[j][5].ToString() == "1")
                            {
                                number1[i] += dt.Rows[j][4].ToString() + "\n";
                                //day[i] += dt.Rows[j][6].ToString() + "\n";
                                //travel[i] += dt.Rows[j][7].ToString() + "\n";
                            }
                            if (dt.Rows[j][1].ToString() == teacher_id[i] && dt.Rows[j][5].ToString() == "2")
                            {
                                number2[i] += dt.Rows[j][4].ToString() + "\n";
                                day[i] += dt.Rows[j][6].ToString() + "\n";
                                travel[i] += dt.Rows[j][7].ToString() + "\n";
                            }

                            //if ((dt.Rows[j][1].ToString() == teacher_id[i] && dt.Rows[j][5].ToString() == "1") || (dt.Rows[j][1].ToString() == teacher_id[i] && dt.Rows[j][5].ToString() == "2"))
                            //{
                            //    if (dt.Rows[j][5].ToString() == "1")
                            //    {
                            //        number2[i] += dt.Rows[j][4].ToString() + "\n";
                            //        //day[i] += dt.Rows[j][6].ToString() + "\n";
                            //        //travel[i] += dt.Rows[j][7].ToString() + "\n";
                            //    }
                            //    if (dt.Rows[j][5].ToString() == "2")
                            //    {
                            //        number1[i] += dt.Rows[j][4].ToString() + "\n";
                            //        //day[i] += dt.Rows[j][6].ToString() + "\n";
                            //        //travel[i] += dt.Rows[j][7].ToString() + "\n";
                            //    }
                            //    if (day[i] == null && travel[i] == null)
                            //    {
                            //        day[i] += dt.Rows[j][6].ToString() + "\n";
                            //        travel[i] += dt.Rows[j][7].ToString() + "\n";
                            //    }
                            //}
                        }
                        reqRow = req.NewRow();
                        reqRow[0] = teacher_name[i];
                        reqRow[1] = number1[i];
                        reqRow[2] = number2[i];
                        reqRow[3] = day[i];
                        reqRow[4] = travel[i];
                        req.Rows.Add(reqRow);

                    }
                    ds.Tables.Add(req);
                }

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;

                return View(ds);
            }
        }

        [HttpPost]
        public ActionResult erty(match TandC_id)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                int count_row = TandC_id.matching.Length;  /////////// นับแถวของ ไอดี ที่เข้ามา

                string[] company_id = new string[count_row];
                string[] teacher_id = new string[count_row];

                DataTable dt;
                DataTable dt1;
                DataTable dt2;

                for (int i = 0; i < count_row; i++)
                {
                    string[] split_id = TandC_id.matching[i].Split('-');
                    company_id[i] = split_id[1];
                    teacher_id[i] = split_id[0];

                    var sql = "select * from Student s inner join Company c on s.company_id = c.company_id inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and c.company_id = '" + company_id[i] + "' and s.semester_out = '" + SEMESTER_ID + "' and s.confirm_status = 'รับแล้ว' order by c.company_id; ";
                    sql += "select distinct company_id from `match` where teacher_id = '" + teacher_id[i] + "';";
                    sql += "select * from `match`";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");
                    dt = ds.Tables[0];
                    dt1 = ds.Tables[1];
                    dt2 = ds.Tables[2];
                    conn.Close();

                    for (int j = 0; j < dt.Rows.Count; j++)
                    {
                        var sql1 = "INSERT INTO `match` (match_id, teacher_id, company_id, student_id,  number, date, date_supervision, semester_id) VALUES('" + (dt2.Rows.Count + j + 1).ToString() + "','" + teacher_id[i] + "','" + company_id[i] + "','" + dt.Rows[j]["student_id"].ToString() + "','1','ยังไม่ได้กำหนด','ยังไม่ได้กำหนด','" + SEMESTER_ID+"');";

                        conn.Open();
                        MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                        DataSet ds1 = new DataSet();
                        da1.Fill(ds1, "TeachersandAuthorities");
                        conn.Close();
                    }
                }
            }

            return RedirectToAction("Teachermatching", "TeachersAndAuthorities");
        }

        public ActionResult teacher()
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (TA_POSITION == "ผู้ช่วยอาจารย์ประจำภาควิชา")
                {
                    /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                    var sql = "SELECT t.teacher_id, t.name, a.password, t.status FROM TeachersandAuthorities t inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and a.position = 'อาจารย์นิเทศ'; ";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");
                    DataTable dt = ds.Tables["TeachersandAuthorities"];
                    conn.Close();

                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["name"] = TA_NAME;
                    ViewData["profile_id"] = TA_PROFILE;
                    ViewData["department"] = TA_MAJOR;
                    ViewData["semCheck"] = SEMESTER_ID;

                    return View(dt);
                }
                if (TA_POSITION == "อาจารย์ฝ่ายสหกิจศึกษา")
                {
                    /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                    var sql = "SELECT t.teacher_id, t.name, a.password, t.status FROM TeachersandAuthorities t inner join Account a on a.user_id = t.teacher_id where a.major = '" + TA_MAJOR + "' and a.position != 'อาจารย์ฝ่ายสหกิจศึกษา' order by a.position;";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");
                    DataTable dt = ds.Tables["TeachersandAuthorities"];
                    conn.Close();

                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["name"] = TA_NAME;
                    ViewData["profile_id"] = TA_PROFILE;
                    ViewData["department"] = TA_MAJOR;
                    ViewData["semCheck"] = SEMESTER_ID;

                    return View(dt);
                }
                return null;
            }
        }

        public ActionResult teacherEditAccount(string id)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select t.teacher_id, t.name, a.password, a.major, t.status from TeachersandAuthorities t inner join Account a on t.teacher_id = a.user_id where a.major = '"+TA_MAJOR+"' and a.user_id = '"+id+"';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables["TeachersandAuthorities"];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;

                ViewData["id"] = dt.Rows[0][0].ToString();
                ViewData["name"] = dt.Rows[0][1].ToString();
                ViewData["pwd"] = dt.Rows[0][2].ToString();
                ViewData["status"] = dt.Rows[0][4].ToString();

                ViewData["t-name"] = TA_NAME;
                ViewData["t-profile_id"] = TA_PROFILE;
                ViewData["t-department"] = TA_MAJOR;
                ViewData["semCheck"] = SEMESTER_ID;

                return View();
            }
        }

        [HttpPost]
        public ActionResult teacherEditAccount(Teachers_And_Authorities ta, string id)
        {
            string UpdateDate = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (ta.pwd == null || ta.pwd == "" || ta.teacher_id == "")
                {
                    return RedirectToAction("teacherEditAccount", "TeachersAndAuthorities");
                }
                else
                {
                    string sql = @"UPDATE TeachersandAuthorities set name = '" + ta.name + @"',status = '" + ta.status + @"' where teacher_id = '" + id + "'; UPDATE Account set password = '" + ta.pwd + @"', date_time_update = '" + UpdateDate + @"' where user_id = '" + id + "';";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "StudentDb");
                    DataTable dt = ds.Tables["StudentDb"];
                    conn.Close();

                    ViewData["teacher_id"] = TA_ID;

                    return RedirectToAction("Teacher", "TeachersAndAuthorities");
                }
            }
        }

        public ActionResult teacherAddAccount()
        {
            ViewData["teacher_id"] = TA_ID;
            //ViewData["major"] = TA_MAJOR;
            ViewData["name"] = TA_NAME;
            ViewData["profile_id"] = TA_PROFILE;
            ViewData["department"] = TA_MAJOR;
            ViewData["semCheck"] = SEMESTER_ID;
            return View();
        }

        [HttpPost]
        public ActionResult teacherAddAccount(Teachers_And_Authorities ta)
        {
            string StartDate = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql1 = "select teacher_id from TeachersandAuthorities where teacher_id = '" + ta.teacher_id + "'; select user_id from Account where user_id = '" + ta.teacher_id + "';";
                conn.Open();
                MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                DataSet ds1 = new DataSet();
                da1.Fill(ds1, "TeachersandAuthorities");
                DataTable dt1 = ds1.Tables[0];
                DataTable dt2 = ds1.Tables[1];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;

                if ((dt1.Rows.Count == 0 && ta.teacher_id != null && ta.pwd != null && ta.name != null) && dt2.Rows.Count == 0)
                {
                    //string sql = @"INSERT INTO TeachersandAuthorities (teacher_id, name, lastname, department, position, pwd, status) VALUES('" + ta.teacher_id + "','" + ta.name + "','','" + TA_MAJOR + "','" + ta.position + "','" + ta.pwd + "','" + ta.status + "')";
                    //string sql2 = @"INSERT INTO Account (user_id, password, position, status, date_time_in, date_time_update) VALUES('" + ta.teacher_id + "','" + ta.pwd + "','" + ta.position + "','ยังไม่มี','" + StartDate + "','')";
                    string sql = @"INSERT INTO TeachersandAuthorities (teacher_id, name, status) VALUES('" + ta.teacher_id + "','" + ta.name + "','" + ta.status + "')";
                    string sql2 = @"INSERT INTO Account (user_id, password, position, major, date_time_in) VALUES('" + ta.teacher_id + "','" + ta.pwd + "','" + ta.position + "','"+TA_MAJOR+ "','" + StartDate + "')";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");
                    conn.Close();

                    conn.Open();
                    MySqlDataAdapter da2 = new MySqlDataAdapter(sql2, conn);
                    DataSet ds2 = new DataSet();
                    da2.Fill(ds2, "Account");
                    conn.Close();

                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["name"] = TA_NAME;
                    ViewData["profile_id"] = TA_PROFILE;
                    ViewData["department"] = TA_MAJOR;
                    ViewData["semCheck"] = SEMESTER_ID;
                    ViewData["notification"] = null;

                    return RedirectToAction("Teacher", "TeachersAndAuthorities");
                }
                else
                {
                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["name"] = TA_NAME;
                    ViewData["profile_id"] = TA_PROFILE;
                    ViewData["department"] = TA_MAJOR;
                    ViewData["semCheck"] = SEMESTER_ID;
                    ViewData["notification"] = "ไม่ได้";
                    return View();
                }
            }
        }

        public ActionResult deleteTeacherAccount(string id)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "DELETE FROM TeachersandAuthorities WHERE teacher_id = '" + id + "';DELETE FROM Account WHERE user_id = '" + id + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables["TeachersandAuthorities"];
                conn.Close();
                return RedirectToAction("teacher", "TeachersAndAuthorities");
            }
        }

        public ActionResult changDatePage(string TandC_id)
        {
            string[] split_id = TandC_id.Split('-');
            string company_id = split_id[0];
            string teacher_id = split_id[1];

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                /// แก้แล้ว/var sql = "select distinct c.company_id, c.name_th, t.teacher_id, t.name, m.date from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id where s.major = '" + TA_MAJOR + "' and t.department = '" + TA_MAJOR + "' and t.status = 'พร้อมออกนิเทศ' and c.company_id = '" + company_id + "' and t.teacher_id = '" + teacher_id + "';";
                var sql = "select distinct c.company_id, c.name_th, t.teacher_id, t.name, m.date from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '"+TA_MAJOR+"' and a2.major = '"+TA_MAJOR+"' and t.status = 'พร้อมออกนิเทศ' and c.company_id = '"+company_id+"' and t.teacher_id = '"+teacher_id+"'; ";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;

                ViewData["company_id"] = dt.Rows[0][0].ToString();
                ViewData["company_name"] = dt.Rows[0][1].ToString();
                ViewData["teacher_id"] = dt.Rows[0][2].ToString();
                ViewData["date"] = dt.Rows[0][4].ToString();

                //ViewData["date"] = DateTime.Now.ToString("MM-dd-yyyy", new CultureInfo("en-US"));
                //string ui = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

                return View();

            }
        }

        [HttpPost]
        public ActionResult changDatePage(chang_date_teacher TCDate)
        {           
            int convert_year = TCDate.datepick.Year + (543 * 2); // แปลงจาก 2019 - 2562
            string date = TCDate.datepick.Day + "/" + TCDate.datepick.Month + "/" + convert_year;
            
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                string sql = @"UPDATE `match` set date = '" + date + @"',date_supervision = '" + date + @"' where teacher_id = '" + TCDate.teacher_id + "' and company_id = '" + TCDate.company_id + "' and semester_id = '" + SEMESTER_ID + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "`match`");
                conn.Close();

                ViewData["teacher_id"] = TA_ID;

            }

            return RedirectToAction("Teachermatching", "TeachersAndAuthorities");
        }

        public ActionResult Student(string semster)
        {
            //using (MySqlConnection conn = new MySqlConnection(strConnection))
            //{
            //    if (semster == null)
            //    {
            //        if (SEMESTER_ID == null)
            //        {
            //            DataSet ds = new DataSet();
            //            ViewData["teacher_id"] = TA_ID;
            //            //ViewData["major"] = TA_MAJOR;
            //            ViewData["term"] = "ช่วงปิดเทอม";
            //            ViewData["semester"] = "ช่วงปิดเทอม";

            //            ViewData["name"] = TA_NAME;
            //            ViewData["profile_id"] = TA_PROFILE;
            //            ViewData["department"] = TA_MAJOR;
            //            ViewData["semCheck"] = SEMESTER_ID;

            //            return View(ds);
            //        }
            //        else
            //        {
            //            string[] split_id = SEMESTER_ID.Split('-');
            //            string year = split_id[0];
            //            string term = split_id[1];

            //            var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and s.semester_out = '" + SEMESTER_ID + "' and s.company_id is not null;";
            //            sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out from Student s inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and s.semester_id = '" + SEMESTER_ID + "' and s.semester_id != s.semester_out;";
            //            sql += "select * from DefineSemester order by semester_id desc;";

            //            conn.Open();
            //            MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
            //            DataSet ds = new DataSet();
            //            da.Fill(ds, "TeachersandAuthorities");

            //            conn.Close();

            //            ViewData["teacher_id"] = TA_ID;
            //            //ViewData["major"] = TA_MAJOR;
            //            ViewData["term"] = "เทอม " + term + " ปีการศึกษา " + year;
            //            ViewData["semester"] = SEMESTER_ID;

            //            ViewData["name"] = TA_NAME;
            //            ViewData["profile_id"] = TA_PROFILE;
            //            ViewData["department"] = TA_MAJOR;
            //            ViewData["semCheck"] = SEMESTER_ID;

            //            return View(ds);

            //        }


            //    }
            //    else
            //        {
            //            string[] split_id = semster.Split('-');
            //        string year = split_id[0];
            //        string term = split_id[1];

            //        var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and s.semester_out = '" + semster + "' and s.company_id is not null;";
            //        sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out from Student s inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and s.semester_id = '" + semster + "' and s.semester_id != s.semester_out;";
            //        sql += "select * from DefineSemester order by semester_id desc;";

            //        conn.Open();
            //        MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
            //        DataSet ds = new DataSet();
            //        da.Fill(ds, "TeachersandAuthorities");

            //        conn.Close();

            //        ViewData["teacher_id"] = TA_ID;
            //        ViewData["major"] = TA_MAJOR;
            //        ViewData["teacher_id"] = TA_ID;
            //        ViewData["major"] = TA_MAJOR;
            //        ViewData["term"] = "เทอม " + term + " ปีการศึกษา " + year;
            //        ViewData["semester"] = semster;

            //        ViewData["name"] = TA_NAME;
            //        ViewData["profile_id"] = TA_PROFILE;
            //        ViewData["department"] = TA_MAJOR;
            //        ViewData["semCheck"] = SEMESTER_ID;

            //        return View(ds);
            //    }

            //}

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                if (semster == null)
                {
                    if (SEMESTER_ID == null)
                    {
                        var sql = "select * from DefineSemester order by semester_id desc;";

                        conn.Open();
                        MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "TeachersandAuthorities");

                        conn.Close();

                        ViewData["teacher_id"] = TA_ID;
                        //ViewData["major"] = TA_MAJOR;
                        ViewData["term"] = "ช่วงปิดเทอม";
                        ViewData["semester"] = "ช่วงปิดเทอม";

                        ViewData["name"] = TA_NAME;
                        ViewData["profile_id"] = TA_PROFILE;
                        ViewData["department"] = TA_MAJOR;
                        ViewData["semCheck"] = SEMESTER_ID;

                        return View(ds);
                    }
                    else
                    {
                        string[] split_id = SEMESTER_ID.Split('-');
                        string year = split_id[0];
                        string term = split_id[1];

                        var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and s.semester_out = '" + SEMESTER_ID + "' and s.company_id is not null;";
                        sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out from Student s inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and s.semester_id = '" + SEMESTER_ID + "' and s.semester_id != s.semester_out;";
                        sql += "select * from DefineSemester order by semester_id desc;";

                        conn.Open();
                        MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "TeachersandAuthorities");

                        conn.Close();

                        ViewData["teacher_id"] = TA_ID;
                        //ViewData["major"] = TA_MAJOR;
                        ViewData["term"] = "เทอม " + term + " ปีการศึกษา " + year;
                        ViewData["semester"] = SEMESTER_ID;

                        ViewData["name"] = TA_NAME;
                        ViewData["profile_id"] = TA_PROFILE;
                        ViewData["department"] = TA_MAJOR;
                        ViewData["semCheck"] = SEMESTER_ID;

                        return View(ds);

                    }


                }
                else
                {
                    string[] split_id = semster.Split('-');
                    string year = split_id[0];
                    string term = split_id[1];

                    var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and s.semester_out = '" + semster + "' and s.company_id is not null;";
                    sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out from Student s inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and s.semester_id = '" + semster + "' and s.semester_id != s.semester_out;";
                    sql += "select * from DefineSemester order by semester_id desc;";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "TeachersandAuthorities");

                    conn.Close();

                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["teacher_id"] = TA_ID;
                    ViewData["major"] = TA_MAJOR;
                    ViewData["term"] = "เทอม " + term + " ปีการศึกษา " + year;
                    ViewData["semester"] = semster;

                    ViewData["name"] = TA_NAME;
                    ViewData["profile_id"] = TA_PROFILE;
                    ViewData["department"] = TA_MAJOR;
                    ViewData["semCheck"] = SEMESTER_ID;

                    return View(ds);
                }

            }

        }

        [HttpPost]
        public ActionResult student(Define_Semester sem) 
        {           
            return RedirectToAction("student", "TeachersAndAuthorities", new { semster = sem.semester_id });
        }

        public ActionResult Requirement_Set_Date() /// กำหนดวันส่งแบบสอบถาม
        {
            //DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));
            //string d = nowdate.Year + "-0" + nowdate.Month + "-0" + nowdate.Day;
            //DateTime startdate = Convert.ToDateTime("");
            //DateTime enddate = Convert.ToDateTime("");
            //TimeSpan StayDay1 = startdate - nowdate;
            //TimeSpan StayDay2 = enddate - nowdate;

            //int DayStay1 = Convert.ToInt16(StayDay1.Days);
            //int DayStay2 = Convert.ToInt16(StayDay2.Days);

            //if (DayStay1 <= 0 && DayStay2 >= 0)
            //{
            //    SEMESTER_ID = "";
            //}

            ViewData["teacher_id"] = TA_ID;
            ViewData["name"] = TA_NAME;
            ViewData["profile_id"] = TA_PROFILE;
            ViewData["department"] = TA_MAJOR;

            ViewData["date_start"] = "2018-03-22";
            ViewData["date_end"] = "2018-02-22";

            return View();
        }

        [HttpPost]
        public ActionResult Requirement_Set_Date(chang_date_teacher date)  /// กำหนดวันส่งแบบสอบถาม
        {
            string date_end = date.datepick.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));
            string date_start = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            DateTime end_date_convert = Convert.ToDateTime(date_end); // แปลงจาก 2019 - 2562
            int convert_date = end_date_convert.Year + (543 * 2); // แปลงจาก 2019 - 2562
            string date_end_finish = end_date_convert.Day + "/" + end_date_convert.Month + "/" + convert_date.ToString(); // แปลงจาก 2019 - 2562

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                ///// แก้ไขแล้ว /var sql2 = "select * from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id where t.department = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "'";
                var sql2 = "select * from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on a.user_id = t.teacher_id where a.major = '"+TA_MAJOR+"' and r.semester_id = '"+SEMESTER_ID+"';";

                conn.Open();
                MySqlDataAdapter da2 = new MySqlDataAdapter(sql2, conn);
                DataSet ds2 = new DataSet();
                da2.Fill(ds2, "RequirementSetDate");
                DataTable dt2 = ds2.Tables[0];
                conn.Close();

                if (dt2.Rows.Count == 0)
                {
                    var sql = "select setdate_id from RequirementSetDate";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "RequirementSetDate");
                    DataTable dt = ds.Tables[0];
                    conn.Close();

                    var sql1 = "INSERT INTO RequirementSetDate (setdate_id, teacher_id, semester_id, startdate, enddate) VALUES('" + (dt.Rows.Count + 1).ToString() + "','" + TA_ID + "','" + SEMESTER_ID + "','" + date_start + "','" + date_end_finish + "')";

                    conn.Open();
                    MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "RequirementSetDate");
                    conn.Close();

                }
                else // if (dt2.Rows.Count == 1 && dt2.Rows[0][1].ToString() == TA_ID) ตรวจว่าคนเปลี่ยนวันที่เป็นคนเดิมหรือไม่
                {
                    //string sql = @"UPDATE RequirementSetDate set enddate = '"+date_end_finish+"' where teacher_id = '"+TA_ID+"' and semester_id = '"+SEMESTER_ID+"'";

                    string sql = @"UPDATE RequirementSetDate set enddate = '" + date_end_finish + "' where teacher_id in (select r.teacher_id from RequirementSetDate r inner join TeachersandAuthorities t on r.teacher_id = t.teacher_id inner join Account a on t.teacher_id = a.user_id where a.major = '" + TA_MAJOR + "' and r.semester_id = '" + SEMESTER_ID + "');";

                    conn.Open();
                    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "`match`");
                    conn.Close();              
                }
            }
            ViewData["teacher_id"] = TA_ID;

            return RedirectToAction("TeacherHome", "TeachersAndAuthorities", new { id = TA_ID });
        }

        public ActionResult StudentAdd(DataTable csvStudent)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select * from DefineSemester order by semester_id desc;";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");

                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["semCheck"] = SEMESTER_ID;
                conn.Close();

                //if(csvStudent.Rows.Count != 0)
                //{
                //    DataTable csvAdd = new DataTable();
                //    csvAdd.Columns.Add("id", typeof(string));
                //    csvAdd.Columns.Add("name", typeof(string));
                //    csvAdd.Columns.Add("teacher_id", typeof(string));
                //    DataRow csvRow;

                //    for (int i = 0; i < csvStudent.Rows.Count; i++)
                //    {
                //        csvRow = csvAdd.NewRow();
                //        csvRow[0] = csvStudent.Rows[i][0].ToString();
                //        csvRow[1] = csvStudent.Rows[i][1].ToString();
                //        csvRow[2] = csvStudent.Rows[i][2].ToString();
                //        csvAdd.Rows.Add(csvRow);
                //    }

                //    ds.Tables.Add(csvAdd);

                //    return View(ds);
                //}
                //else
                //{
                    
                //}
                return View(ds);
            }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult StudentAdd(HttpPostedFileBase upload, Define_Semester sem)  ///////////////////// Upload .CSV file
        {
            string date_time_in = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            if (ModelState.IsValid)
            {
                if (upload != null && upload.ContentLength > 0)
                {
                    if (upload.FileName.EndsWith(".csv"))
                    {
                        //DataTable dt;
                        Stream stream = upload.InputStream;
                        DataTable csvTable = new DataTable();
                        using (CsvReader csvReader = new CsvReader(new StreamReader(stream), true))
                        {
                            csvTable.Load(csvReader);
                        }

                        //using (MySqlConnection conn = new MySqlConnection(strConnection))
                        //{
                        //    var sql = "select user_id from Account;";

                        //    conn.Open();
                        //    MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                        //    DataSet ds = new DataSet();
                        //    da.Fill(ds, "student");
                        //    conn.Close();

                        //    for (int i = 0; i < csvTable.Rows.Count; i++)
                        //    {
                        //        string[] split_id = csvTable.Rows[i][0].ToString().Split('-');
                        //        string user_id = split_id[0] + split_id[1];
                        //        string password = user_id.Remove(0, 9);

                        //        var sql1 = "INSERT INTO Account (user_id, password, position, major, status, date_time_in, date_time_update) VALUES('"+user_id+"','"+password+"','นักศึกษา','"+TA_MAJOR+"','ยังไม่มี','"+ date_time_in + "','')";
                        //            sql1 += "INSERT INTO Student (student_id, name_th, advisor, semester_id, semester_out) VALUES('" + user_id + "','" + csvTable.Rows[i][1].ToString() + "','"+ csvTable.Rows[i][2].ToString() + "','"+sem.semester_id+ "','"+ sem.semester_id + "')";

                        //        conn.Open();
                        //        MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                        //        DataSet ds1 = new DataSet();
                        //        da1.Fill(ds1, "student");
                        //        conn.Close();
                        //    }
                        //}

                        //return View(/*dt*/);
                        return RedirectToAction("StudentAdd", "TeachersAndAuthorities", new { csvStudent = csvTable });
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        //return RedirectToAction("student", "TeachersAndAuthorities");
                        return View();
                    }
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            //return RedirectToAction("student", "TeachersAndAuthorities");
            return RedirectToAction("StudentAdd", "TeachersAndAuthorities");
        }

        public ActionResult StudentAdd_Confirm()
        {
            ViewData["teacher_id"] = TA_ID;
            ViewData["name"] = TA_NAME;
            ViewData["profile_id"] = TA_PROFILE;
            ViewData["department"] = TA_MAJOR;
            ViewData["semCheck"] = SEMESTER_ID;

            return View();
        }

        [HttpPost]
        public ActionResult StudentAdd_Confirm(HttpPostedFileBase upload, Define_Semester sem)  ///////////////////// Upload .CSV file
        {
            string date_time_in = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            if (ModelState.IsValid)
            {
                if (upload != null && upload.ContentLength > 0)
                {
                    if (upload.FileName.EndsWith(".csv"))
                    {
                        Stream stream = upload.InputStream;
                        DataTable csvTable = new DataTable();

                        using (CsvReader csvReader = new CsvReader(new StreamReader(stream), true))
                        {
                            csvTable.Load(csvReader);
                        
                            ViewData["teacher_id"] = TA_ID;
                            ViewData["name"] = TA_NAME;
                            ViewData["profile_id"] = TA_PROFILE;
                            ViewData["department"] = TA_MAJOR;
                            ViewData["term"] = sem.semester_id;
                            ViewData["semCheck"] = SEMESTER_ID; //////

                            return View(csvTable);
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("File", "This file format is not supported");
                        return View();
                    }
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }

            return RedirectToAction("StudentAdd", "TeachersAndAuthorities");
        }

        [HttpPost]
        public ActionResult ConfirmCsv(Csv csvConfirm, Define_Semester sem)
        {
            string date_time_in = DateTime.Now.ToString("dd/MM/yyyy", new CultureInfo("th-TH"));

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                for (int i = 0; i < csvConfirm.id.Length; i++)
                {
                    string[] split_id = csvConfirm.id[i].Split('-');
                    string user_id = split_id[0] + split_id[1];
                    string password = user_id.Remove(0, 9);

                    //var sql1 = "INSERT INTO Account (user_id, password, position, major, status, date_time_in, date_time_update) VALUES('" + user_id + "','" + password + "','นักศึกษา','" + TA_MAJOR + "','ยังไม่มี','" + date_time_in + "','')";
                    //sql1 += "INSERT INTO Student (student_id, name_th, advisor, semester_id, semester_out) VALUES('" + user_id + "','" + csvConfirm.name[i].ToString() + "','" + csvConfirm.teacher[i].ToString() + "','" + sem.semester_id + "','" + sem.semester_id + "')";

                    var sql1 = "INSERT INTO Account (user_id, password, position, major, date_time_in) VALUES('" + user_id + "','" + password + "','นักศึกษา','" + TA_MAJOR + "','" + date_time_in + "');";
                    sql1 += "INSERT INTO Student (student_id, name_th, advisor, semester_id, semester_out, sec) VALUES('" + user_id + "','" + csvConfirm.name[i].ToString() + "','" + csvConfirm.teacher[i].ToString() + "','" + sem.semester_id + "','" + sem.semester_id + "','"+ csvConfirm.sec[i].ToString() + "');";
                    sql1 += "INSERT INTO Supervisor (student_id) VALUES('" + user_id + "');";
                    sql1 += "insert into Document(student_id) values('" + user_id + "') ;";

                    conn.Open();
                    MySqlDataAdapter da1 = new MySqlDataAdapter(sql1, conn);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, "student");
                    conn.Close();
                }
            }

            return RedirectToAction("StudentAdd", "TeachersAndAuthorities");
        }

        public ActionResult Accsesment(string s_ID)
        {
            AccesmentScore a = new AccesmentScore();

            string[] split_id = s_ID.Split('+');
            string id = split_id[0];
            string sem = split_id[1];

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                var sql = "select s.student_id, s.name_th, t.name, m.semester_id , d.score16, d.title_th, d.title_en, d.comment, c.name_th from Document d inner join `match` m on d.student_id = m.student_id inner join Student s on s.student_id = m.student_id inner join Company c on c.company_id = s.company_id  inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and m.semester_id = '" + sem + "' and s.student_id = '" + id + "' and d.student_id = '" + id + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();
               
                int score, total = 0;

                string score16 = dt.Rows[0][4].ToString();
                string[] splitScore16 = score16.Split(',');

                for (int k = 0; k < 10; k++)
                {
                    int.TryParse(splitScore16[k], out score);
                    total = total + score;
                }

                a.ds = getDataset(ds);
                a.score = getScore(score16);

                ViewData["teacher_id"] = TA_ID;
                ViewData["major"] = TA_MAJOR;
                ViewData["student_id"] = dt.Rows[0][0].ToString();
                ViewData["student_name"] = dt.Rows[0][1].ToString();
                ViewData["project_name"] = dt.Rows[0][5].ToString();
                ViewData["teacher_name"] = dt.Rows[0][2].ToString();
                ViewData["company_name"] = dt.Rows[0][8].ToString();
                ViewData["major"] = TA_MAJOR;
                ViewData["score"] = total;
                ViewData["sem"] = sem;

                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["semCheck"] = SEMESTER_ID;

                return View(a);

            }
        }

        public DataSet getDataset(DataSet ds)
        {           
            return ds;
        }

        public DataSet getDatabase(string str_sql)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = str_sql;

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();

                return ds;
            }
        }

        public string[] getScore(string score)
        {
            return score.Split(',');
        }

        public string getTerm()
        {
            string term = null;

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select * from DefineSemester order by semester_id desc;";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];

                conn.Close();

                for (int i = 0; i < dt.Rows.Count; i++)  // ตรวจสอบว่าเป็นเทอมไหน ปีไหน
                {
                    DateTime nowdate = Convert.ToDateTime(DateTime.Now.ToString("dd/MM/yyyy"));
                    DateTime startdate = Convert.ToDateTime(dt.Rows[i]["date_start"]);
                    DateTime enddate = Convert.ToDateTime(dt.Rows[i]["date_end"]);
                    TimeSpan StayDay1 = startdate - nowdate;
                    TimeSpan StayDay2 = enddate - nowdate;

                    int DayStay1 = Convert.ToInt16(StayDay1.Days);
                    int DayStay2 = Convert.ToInt16(StayDay2.Days);

                    if (DayStay1 <= 0 && DayStay2 >= 0)
                    {
                        term = dt.Rows[i]["semester_id"].ToString();
                    }
                }
                return term;
            }
        }

        public ActionResult All_Accesment(string semster)
        {
            string year = null;
            string term = null;

            if (semster != null)
            {
                string[] split_id = semster.Split('-');
                year = split_id[0];
                term = split_id[1];
            }

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                var sql = "select s.student_id, s.name_th, c.name_th, t.name , m.semester_id from Document d inner join `match` m on d.student_id = m.student_id inner join Student s on s.student_id = m.student_id inner join Company c on c.company_id = s.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and m.semester_id = '" + semster + "' and d.score16 is not null order by t.name;";
                sql += "select s.student_id, s.name_th from `match` m inner join Student s on m.student_id = s.student_id inner join Account a on a.user_id = s.student_id where m.semester_id = '" + semster + "' and a.major = '" + TA_MAJOR + "';";
                sql += "select * from DefineSemester order by semester_id desc;";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;
                ViewData["major"] = TA_MAJOR;
                ViewData["term_and_year"] = "เทอม " + term + " ปีการศึกษา " + year;
                ViewData["sem"] = semster;

                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["semCheck"] = SEMESTER_ID;

                return View(ds);

            }
        }

        [HttpPost]
        public ActionResult All_Accesment(Define_Semester sem)
        {
            return RedirectToAction("All_Accesment", "TeachersAndAuthorities", new { semster = sem.semester_id });
        }

        public ActionResult sk16(string s_ID) // สก.16
        {
            string[] split_id = s_ID.Split('+');
            string id = split_id[0];
            string sem = split_id[1];

            string fileName = "20160702-15.pdf";
            string sourcePath = Server.MapPath("~/pdf_file");
            string targetPath = Server.MapPath("~/pdf_file/pdf_copy");

            // Use Path class to manipulate file and directory paths.
            string sourceFile = System.IO.Path.Combine(sourcePath, fileName);
            string destFile = System.IO.Path.Combine(targetPath, TA_ID + "-" + id + "-" + fileName);
            System.IO.File.Copy(sourceFile, destFile, true); // Coppy File

            string oldFile = Server.MapPath("~/pdf_file/" + fileName);
            string newFile = Server.MapPath("~/pdf_file/pdf_copy/" + TA_ID + "-" + id + "-" + fileName);

            //////////////////////////////////////////////////////////////////////////////

            using (var reader = new PdfReader(oldFile))
            {
                using (var fileStream = new FileStream(newFile, FileMode.Create, FileAccess.Write))
                {
                    using (MySqlConnection conn = new MySqlConnection(strConnection))
                    {
                        /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                        //var sql = "select s.student_id, s.name_th, t.name, m.semester_id , d.*, c.name_th, a1.major from Document d inner join `match` m on d.student_id = m.student_id inner join Student s on s.student_id = m.student_id inner join Company c on c.company_id = s.company_id  inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and m.semester_id = '" + SEMESTER_ID + "' and s.student_id = '" + s_ID + "' and d.student_id = '" + s_ID + "';";

                        var sql = "select s.student_id, s.name_th, t.name, m.semester_id , d.score16, d.title_th, d.title_en, d.comment, c.name_th, a1.major from Document d inner join `match` m on d.student_id = m.student_id inner join Student s on s.student_id = m.student_id inner join Company c on c.company_id = s.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = t.teacher_id inner join Account a2 on a2.user_id = s.student_id where a1.major = '"+TA_MAJOR+"' and a2.major = '"+TA_MAJOR+"' and m.semester_id = '"+ sem + "' and s.student_id = '"+ id + "' and d.student_id = '"+ id + "';";

                        conn.Open();
                        MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                        DataSet ds = new DataSet();
                        da.Fill(ds, "TeachersandAuthorities");
                        DataTable dt = ds.Tables[0];
                        conn.Close();

                        int score, total = 0;

                        string score16 = dt.Rows[0][4].ToString();
                        string[] splitScore16 = score16.Split(',');

                        for (int k = 0; k < 10; k++)
                        {
                            int.TryParse(splitScore16[k], out score);
                            total = total + score;
                        }

                        ////////////////////////////////////////////////////////////////////////////

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

                                cb.AddTemplate(importedPage, 0, 0);

                                cb.BeginText();
                                text = dt.Rows[0][1].ToString(); // "นาย เนติพงษ์ สุพัฒสร";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 225, 570, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][0].ToString(); //"116030462031-4";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 450, 570, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][9].ToString(); //"วิศวกรรมคอมพิวเตอร์";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 218, 552, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = "วิศวกรรมศาสตร์";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 448, 552, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][8].ToString(); //"บริษัท ทีม พรีซิชัน จำกัด";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 155, 533, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][2].ToString(); //"ดร. สิตั๋น ฉันนะ";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 175, 515, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][5].ToString(); //"ไม่รุ้ทำไร";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 140, 459, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][6].ToString(); //"WHat do you do?";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, 165, 442, 0);
                                cb.EndText();
                                //////////////////////////////////////////////////////////////////

                                cb.BeginText();
                                text = splitScore16[0]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 375, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[1]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 339, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[2]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 303, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[3]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 268, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[4]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 231, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[5]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 193, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[6]; //"7";
                                cb.ShowTextAligned(PdfContentByte.ALIGN_CENTER, text, 502, 158, 0);
                                cb.EndText();

                                //cb.AddTemplate(importedPage, 0, 0);
                            }

                            if (i == 2)
                            {
                                document.NewPage();

                                BaseFont baseFont = BaseFont.CreateFont(Server.MapPath("~/fonts/THSarabunNew.ttf"), BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                                PdfImportedPage importedPage = writer.GetImportedPage(reader, i);

                                PdfContentByte cb = writer.DirectContent;
                                cb.SetColorFill(BaseColor.BLACK);
                                cb.SetFontAndSize(baseFont, 12);

                                cb.AddTemplate(importedPage, 0, 0);

                                string text;

                                cb.BeginText();
                                text = splitScore16[7]; //"7";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_LEFT, text, 495, 711, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[8]; //"7";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_LEFT, text, 495, 674, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = splitScore16[9].ToString(); //"7";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_LEFT, text, 495, 638, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][7].ToString(); //"Example";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_LEFT, text, 90, 566, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = dt.Rows[0][2].ToString().ToString(); //"อาจารย์";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_CENTER, text, 429, 466, 0);
                                cb.EndText();

                                cb.BeginText();
                                text = total.ToString(); //"รวม 100";
                                cb.ShowTextAlignedKerned(PdfContentByte.ALIGN_LEFT, text, 456, 336, 0);
                                cb.EndText();

                            }
                        }

                        document.Close();
                        writer.Close();
                    }
                }
            }
            return RedirectToAction("GetReportPdf", "TeachersAndAuthorities", new { pdf_id = TA_ID + "-" + id + "-" + fileName });
        }

        public FileResult GetReportPdf(string pdf_id) // เปิด PDF ในหน้าใหม่
        {
            string ReportURL = Server.MapPath("~/pdf_file/pdf_copy/" + pdf_id);
            byte[] FileBytes = System.IO.File.ReadAllBytes(ReportURL);

            string physicalPath = Server.MapPath("~/pdf_file/pdf_copy/" + pdf_id);
            var uri = new Uri(physicalPath, UriKind.Absolute); //ลบไฟล์
            System.IO.File.Delete(uri.LocalPath);

            return File(FileBytes, "application/pdf");
        }

        public ActionResult DeleteFile(string pdf_id) // ลบ file PDF Copy
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                /// แก้ไขแล้ว///var sql = "SELECT * FROM TeachersandAuthorities where department = '" + TA_MAJOR + "' and position = 'อาจารย์นิเทศ'";
                var sql = "select s.student_id from Student s inner join Account a on s.student_id = a.user_id where a.major = '"+TA_MAJOR+"';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                conn.Close();

                ViewData["teacher_id"] = TA_ID;
                ViewData["major"] = TA_MAJOR;

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    string physicalPath = Server.MapPath("~/pdf_file/pdf_copy/" + TA_ID + "-" + dt.Rows[i][0].ToString() + "-" + "20160702-15.pdf");

                    var uri = new Uri(physicalPath, UriKind.Absolute); //ลบไฟล์
                    System.IO.File.Delete(uri.LocalPath);
                }

                return RedirectToAction("upload_image", "TeachersAndAuthorities");
            }
        }

        public ActionResult excel(string sem)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '"+TA_MAJOR+"' and(s.semester_id = '"+sem+"' or s.semester_out = '"+sem+"');";
                    sql += "select company_id,name_th from Company;";

                sql += "select c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.postal_code, p.province, c.phone, c.fax, s.name_th, a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '"+TA_MAJOR+"' and s.semester_out = '"+sem+"' order by c.name_th;";
                sql += "select  distinct c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.province, p.postal_code, c.phone, c.fax from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem+"' order by c.name_th;";
                sql += "select c.name_th , ct.name, ct.position, ct.department from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id inner join Contact ct on ct.student_company_id = c.company_id where a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem+"' order by c.name_th; ";

                sql += "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '"+TA_MAJOR+"' and s.semester_out = '" + sem + "';";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                DataTable dt_company = ds.Tables[1]; ////

                DataTable dt0 = ds.Tables[2];
                DataTable dt1 = ds.Tables[3];
                DataTable dt2 = ds.Tables[4];

                conn.Close();

                string[] split_sem= sem.Split('-');
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

                        for (int a = 1; a <= 1; a++)
                        {
                            for (int i = 0; i < dt.Rows.Count; i++)
                            {
                                //Merge 
                                if (i == 0)
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
                                    worksheet.Range["B" + (total + 3)].Text = "หลักสูตร 2004060206  :  " + TA_MAJOR;
                                    worksheet.Range["B" + (total + 4)].Text = "ปริญญาตรี";

                                    worksheet.Range["K" + (total + 1)].Text = "รายชื่อนักศึกษา";
                                    worksheet.Range["K" + (total + 2)].Text = "ภาคการศึกษาที่ " + term + "/" + year;
                                    worksheet.Range["K" + (total + 3)].Text = "กลุ่ม 59146 ELE";
                                    worksheet.Range["K" + (total + 4)].Text = "ภาคปกติ  จำนวน " + ds.Tables[5].Rows.Count + " คน";

                                    worksheet.Range["A" + (total + 5)].Text = "ลำดับ";
                                    worksheet.Range["B" + (total + 5)].Text = "รหัสนักศึกษา";
                                    worksheet.Range["C" + (total + 5)].Text = "ชื่อ - สกุล";
                                    worksheet.Range["D" + (total + 5)].Text = "สถานประกอบการ";

                                    worksheet.Range["C" + (total + 4)].Text = "อ. ที่ปรึกษา " + "อาจารย์มาโนชย์ ประชา";
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

                                    worksheet.Range["A" + (total + 6)].Text = (i + 1).ToString();
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

                                    worksheet.Range["A" + (total + 6)].Text = (i + 1).ToString();
                                    worksheet.Range["B" + (total + 6)].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                                    worksheet.Range["B" + (total + 6)].Text = dt.Rows[i][0].ToString();
                                    worksheet.Range["C" + (total + 6)].Text = dt.Rows[i][1].ToString();
                                    worksheet.Range["D" + (total + 6)].Text = "ไม่ออกฝึก";

                                    worksheet.Range["A" + (total + 6) + ":" + "D" + (total + 6)].CellStyle.Font.Color = ExcelKnownColors.Red;

                                    total++;
                                }
                            }
                            total += 10;
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
                            worksheet1.Range["A" + total].Text = (i + 1).ToString();
                            worksheet1.Range["B" + total].Text = dt1.Rows[i][0].ToString();
                            worksheet1.Range["C" + total].Text = dt1.Rows[i][1].ToString() + " " + dt1.Rows[i][2].ToString() + " " + dt1.Rows[i][3].ToString() + " " + dt1.Rows[i][4].ToString() + " " + dt1.Rows[i][5].ToString();
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

                            /////////////////////////////////////
                            //for (int supervise = 0; supervise < dt2.Rows.Count; supervise++)
                            //{
                            //    if (dt1.Rows[i][0].ToString() == dt2.Rows[supervise][0].ToString())
                            //    {
                            //        worksheet1.Range["D" + (total)].Text = dt2.Rows[supervise][1].ToString();
                            //        worksheet1.Range["D" + (total + 1)].Text = dt2.Rows[supervise][2].ToString();
                            //        //total++;
                            //    }
                            //}
                            ////////////////////////////////////

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

                        worksheet1.Range["A1"].Text = "ข้อมูลสถานประกอบการที่ส่งนักศึกษาสหกิจศึกษา ภาคเรียนที่ "+term+"/" +year+ " เข้าฝึกประสบการณ์วิชาชีพ ภาควิชา"+TA_MAJOR+" จำนวน " + total_student + " คน";

                    }

                    worksheet.Name = "รายชื่อนักศึกษา "+sem;
                    worksheet1.Name = "รายชื่อบริษัท " + sem;
                    //Save the workbook to disk in xlsx format
                    workbook.SaveAs("รายงานการออกฝึกงานภาควิชา" +TA_MAJOR+ " ปี " +sem+ ".xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
                }
            }

            return null;
        }

        public ActionResult excel1(string sem)
        {
            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                var sql = "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out, s.sec, s.advisor from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + TA_MAJOR + "' and(s.semester_id = '" + sem + "' or s.semester_out = '" + sem + "');";
                sql += "select company_id,name_th from Company;";

                sql += "select c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.postal_code, p.province, c.phone, c.fax, s.name_th, a.major from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem + "' order by c.name_th;";
                sql += "select  distinct c.name_th, ad.number, ad.lane, ad.road, ad.building, ad.floor, p.sub_area, p.area, p.province, p.postal_code, c.phone, c.fax from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id where a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem + "' order by c.name_th;";
                sql += "select c.name_th , ct.name, ct.position, ct.department from Student s inner join Company c on s.company_id = c.company_id inner join Account a on s.student_id = a.user_id inner join Address ad on ad.address_id = c.address_id inner join PrototypeAddress p on ad.prototype_id = p.prototype_id inner join Contact ct on ct.student_company_id like c.company_id where a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem + "' order by c.name_th; ";

                sql += "select s.student_id,s.name_th,s.company_id,s.semester_id,s.semester_out from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + TA_MAJOR + "' and s.semester_out = '" + sem + "';";

                sql += "select distinct s.sec from Student s inner join Account a on s.student_id = a.user_id where s.student_id = a.user_id and a.major = '" + TA_MAJOR + "' and(s.semester_id = '" + sem + "' or s.semester_out = '" + sem + "')";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                DataTable dt = ds.Tables[0];
                DataTable dt_company = ds.Tables[1]; ////

                DataTable dt0 = ds.Tables[2];
                DataTable dt1 = ds.Tables[3];
                DataTable dt2 = ds.Tables[4];

                conn.Close();

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
                                        worksheet.Range["B" + (total + 3)].Text = "หลักสูตร 2004060206  :  " + TA_MAJOR;
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

                            /////////////////////////////////////
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
                                if(dt1.Rows[i][4].ToString() == null || dt1.Rows[i][4].ToString() == "")
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

                        worksheet1.Range["A1"].Text = "ข้อมูลสถานประกอบการที่ส่งนักศึกษาสหกิจศึกษา ภาคเรียนที่ " + term + "/" + year + " เข้าฝึกประสบการณ์วิชาชีพ ภาควิชา" + TA_MAJOR + " จำนวน " + total_student + " คน";

                    }

                    worksheet.Name = "รายชื่อนักศึกษา " + sem;
                    worksheet1.Name = "รายชื่อบริษัท " + sem;
                    //Save the workbook to disk in xlsx format
                    workbook.SaveAs("รายงานการออกฝึกงานภาควิชา" + TA_MAJOR + " ปี " + sem + ".xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
                }
            }

            return null;
        }

        [HttpPost]
        public ActionResult upload_image(HttpPostedFileBase file)
        {
            if (file != null)
            {
                string ImageName = TA_ID + ".jpg";

                string physicalPath = Server.MapPath("~/assets/images/users/" + ImageName);

                file.SaveAs(physicalPath); // save image in folder

                //var uri = new Uri(physicalPath, UriKind.Absolute); //ลบไฟล์รุป
                //System.IO.File.Delete(uri.LocalPath);

            }
            return RedirectToAction("TeacherEdit", "TeachersAndAuthorities", new { id = TA_ID });
        }

        public ActionResult StudentAll(string semster)
        {
            string year = null;
            string term = null;

            if (semster != null)
            {
                string[] split_id = semster.Split('-');
                year = split_id[0];
                term = split_id[1];
            }

            using (MySqlConnection conn = new MySqlConnection(strConnection))
            {
                //var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out,a1.status from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = 'วิศวกรรมคอมพิวเตอร์' and a2.major = 'วิศวกรรมคอมพิวเตอร์' and(s.semester_out = '"+ semster + "' or s.semester_id = '"+ semster + "') order by s.semester_id; ";
                //sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out,a.status from Student s inner join Account a on a.user_id = s.student_id where a.major = 'วิศวกรรมคอมพิวเตอร์' and a.major = 'วิศวกรรมคอมพิวเตอร์' and(s.semester_out = '"+ semster + "' or s.semester_id = '"+ semster + "'); ";
                //sql += "select * from DefineSemester order by semester_id desc;";

                //var sql = "select s.student_id, s.name_th, c.name_th, t.name, s.semester_id, s.semester_out, s.sec from `match` m inner join Student s on m.student_id = s.student_id inner join Company c on c.company_id = m.company_id inner join TeachersandAuthorities t on t.teacher_id = m.teacher_id inner join Account a1 on a1.user_id = s.student_id inner join Account a2 on a2.user_id = t.teacher_id where a1.major = '" + TA_MAJOR + "' and a2.major = '" + TA_MAJOR + "' and(s.semester_out = '" + semster + "' or s.semester_id = '" + semster + "') order by s.semester_id; ";

                var sql = "select s.student_id, s.name_th, c.name_th, s.confirm_status, s.semester_id, s.semester_out, s.sec from Student s inner join Company c on c.company_id = s.company_id inner join Account a1 on a1.user_id = s.student_id where a1.major = '" + TA_MAJOR + "' and(s.semester_out = '" + semster + "' or s.semester_id = '" + semster + "') order by s.semester_id; ";

                sql += "select s.student_id, s.name_th, s.semester_id, s.semester_out, s.sec from Student s inner join Account a on a.user_id = s.student_id where a.major = '" + TA_MAJOR + "' and(s.semester_out = '" + semster + "' or s.semester_id = '" + semster + "');";
                sql += "select * from DefineSemester order by semester_id desc;";
                sql += "select distinct s.sec from Student s inner join Account a on s.student_id = a.user_id where a.major = '" + TA_MAJOR + "' and(s.semester_out = '" + semster + "' or s.semester_id = '" + semster + "')";

                conn.Open();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "TeachersandAuthorities");
                conn.Close();


                ViewData["teacher_id"] = TA_ID;
                ViewData["name"] = TA_NAME;
                ViewData["profile_id"] = TA_PROFILE;
                ViewData["department"] = TA_MAJOR;
                ViewData["term"] = "เทอม "+term+" ปีการศึกษา " +year;
                ViewData["sem"] = semster;
                ViewData["semCheck"] = SEMESTER_ID;

                return View(ds);

            }
        }

        [HttpPost]
        public ActionResult StudentAll(Define_Semester sem)  
        {
            return RedirectToAction("StudentAll", "TeachersAndAuthorities", new { semster = sem.semester_id });
        }

        public ActionResult taLogout()
        {
            TA_ID = null;
            TA_MAJOR = null;
            TA_NAME = null;
            TA_PROFILE = null;
            SEMESTER_ID = null;
            SEM = null;

            return RedirectToAction("LoginTemplate", "Home");
        }


       


    }
}