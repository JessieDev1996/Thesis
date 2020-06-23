using Project_Test1.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Test1.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult LoginTemplate(string loginMsg)
        {         
            ViewData["loginMsg"] = loginMsg;
            Session.Abandon();
            return View();
        }

        [HttpPost]
        public ActionResult LoginTemplate(Account acc)
        {
            ConnectDB db = new ConnectDB();
            var _login = "SELECT * FROM Account where user_id = '"+acc.user_id+"' and password = '"+ acc.password+"'";
            DataSet ds = new DataSet();
            ds = db.select(_login,"Account");
            var a = ds.Tables["Account"].Rows.Count;
            if (a==1)
            {
                acc.position = ds.Tables["Account"].Rows[0]["position"].ToString();
                if (acc.position == "ผู้ดูแล")
                {
                    Session["position"] = acc.position;
                    return RedirectToAction("Index", "Admin", new { user_id_title = acc.user_id});
                }
                else if (acc.position == "เจ้าหน้าที่ฝ่ายสหกิจศึกษา")
                {
                    Session["position"] = acc.position;
                    return RedirectToAction("Index", "Authorities", new { user_id_title = acc.user_id });
                }
                else if (acc.position == "นักศึกษา")
                {
                    Session["position"] = acc.position;
                    return RedirectToAction("Index", "Student", new { user_id_title = acc.user_id });
                }
                else if (acc.position == "อาจารย์นิเทศ")
                {
                    Session["position"] = acc.position;
                    return RedirectToAction("Index", "Teacher", new { user_id_title = acc.user_id });
                }
                else if (acc.position == "อาจารย์ฝ่ายสหกิจศึกษา" || acc.position == "ผู้ช่วยอาจารย์ประจำภาควิชา")
                {
                    Session["position"] = acc.position;
                    return RedirectToAction("TeacherHome", "TeachersAndAuthorities", new { id = acc.user_id });
                }
               
            }
            else if(acc.user_id != null && acc.password != null)
            {
                acc.loginErrMsg = "Invalid";
            }
            return RedirectToAction("LoginTemplate", new { loginMsg = acc.loginErrMsg });
        }
    }
}