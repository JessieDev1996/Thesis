using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Project_Test1.Models
{
    public class Company
    {
        //company
        public int[] company_id { get; set; }
        public int address_id { get; set; }
        public string[] name_en { get; set; }
        public string[] name_th { get; set; }
        public string[] blacklist { get; set; }
        public string phone { get; set; }
        public string fax { get; set; }
        public string email { get; set; }
        public string website { get; set; }
        public string business_type { get; set; }
        public string job_request { get; set; }
        public string job_time { get; set; }
        public string job_day { get; set; }
        public string pay { get; set; }
        public string accommodation { get; set; }
        public string charge_accommodation { get; set; }
        public string bus { get; set; }
        public string charge_bus { get; set; }
        public string welfare { get; set; }
        public string exam { get; set; }
        public string start_time { get; set; }
        public string update_time { get; set; }
        public string all_employee { get; set; }



        //address
        //public int address_id { get; set; }
        public int prototype_id { get; set; }
        public string student_company_id { get; set; }
        public string who { get; set; }
        public string number { get; set; }
        public string building { get; set; }
        public string floor { get; set; }
        public string lane { get; set; }
        public string road { get; set; }
        public string sub_area { get; set; }
        public string area { get; set; }
        public string province { get; set; }

        public string pX { get; set; }
        public string pY { get; set; }
        //public string postal_code { get; set; }


        //contact
        public int contact_id { get; set; }
        //public string student_company_id { get; set; }
        public int con_number { get; set; }
        public string name { get; set; }
        public string position { get; set; }
        public string department { get; set; }
        public string con_phone { get; set; }
        public string con_fax { get; set; }
        public string con_email { get; set; }


        //contact for ceo
        public string name_s { get; set; }
        public string position_s { get; set; }
        public string department_s { get; set; }


        //semester
        public string year { get; set; }
        public int semester_id1 { get; set; }
        public int semester_id2 { get; set; }
        public int semester_id3 { get; set; }




        //dynamic form recieve of semester
        public int rs_id { get; set; }
        //public string semester_id { get; set; }
        //public int company_id { get; set; }
        public string[] major { get; set; }
        public string[] job_description { get; set; }
        public int[] require { get; set; }
        public string[] position_of_student { get; set; }


        //term
        public string[] term { get; set; }

    }
}