using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Project_Test1.Models
{
    public class Student
    {
        public string student_id { get; set; }
        public int company_id { get; set; }
        public int address_id { get; set; }
        public int document_id { get; set; }
        public string semester_id { get; set; }
        [Required(ErrorMessage = "This field is required.")]
        public string name_en { get; set; }
        public string name_th { get; set; }
        public string study { get; set; }
        public string major { get; set; }
        public Nullable<int> year { get; set; }
        public string advisor { get; set; }
        public Nullable<float> gpa { get; set; }
        public Nullable<float> gpax { get; set; }
        public string identification { get; set; }
        public string at { get; set; }
        public string date { get; set; }
        public string telephone { get; set; }
        public string mobile_phone { get; set; }
        public string fax { get; set; }
        public string email { get; set; }
        public string race { get; set; }
        public string nationality { get; set; }
        public string religion { get; set; }
        public string date_birth { get; set; }
        public Nullable<int> age { get; set; }
        public string gender { get; set; }
        public Nullable<int> height { get; set; }
        public Nullable<int> weight { get; set; }
        public string disease { get; set; }
        public int no_relative { get; set; }
        public int number { get; set; }
        public string ability_spectify { get; set; }
        public string experience { get; set; }
        public string confirm_status { get; set; }


        //Languages table
        public int language_id { get; set; }
        public int la_number { get; set; }
        public string[] la_name { get; set; }
        public string[] la_listening { get; set; }
        public string[] la_speaking { get; set; }
        public string[] la_reading { get; set; }
        public string[] la_writing { get; set; }

        //Family table
        public int family_id { get; set; }
        public int fa_number { get; set; }
        public string[] fa_name { get; set; }
        public string[] fa_relation { get; set; }
        public int[] fa_age { get; set; }
        public string[] fa_occupation { get; set; }
        public string[] fa_position { get; set; }
        public string[] fa_telephone { get; set; }
        public string[] fa_mobile_phone { get; set; }
        public string[] fa_fax{ get; set; }
        public string[] fa_email { get; set; }
        public string[] fa_place_work { get; set; }



        //History of Education Table
        public int history_id { get; set; }
        public int hi_number { get; set; }
        public string[] hi_name { get; set; }
        public int[] hi_year { get; set; }
        public string[] hi_certificate { get; set; }
        public string[] hi_major { get; set; }
        public float[] hi_gpa { get; set; }


        //Address Table
        //public string student_company_id { get; set; }
        public string[] ad_who { get; set; }
        public string[] ad_number { get; set; }
        public string[] ad_building { get; set; }
        public string[] ad_floor { get; set; }
        public string[] ad_lane { get; set; }
        public string[] ad_road { get; set; }
        public string[] ad_sub_area { get; set; }
        public string[] ad_area { get; set; }
        public string[] ad_province { get; set; }
        public int[] ad_prototype_id { get; set; }

    }
}