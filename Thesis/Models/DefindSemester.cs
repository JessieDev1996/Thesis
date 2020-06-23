using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace Project_Test1.Models
{
    public class DefineSemester
    {
        [DisplayName("ปีการศึกษา:")]
        public string semester_id { get; set; }
        [DisplayName("วันที่เริ่มต้นสหกิจ/ฝึกงาน:")]
        public string date_start1 { get; set; }
        [DisplayName("วันที่สิ้นสุดสหกิจ/ฝึกงาน:")]
        public string date_end1 { get; set; }
        [DisplayName("จำนวนครั้งในการออกนิเทศ:")]
        [Required(ErrorMessage = "This field is required.")]
        public Nullable<int> to_go1 { get; set; }
        [DisplayName("วันที่เริ่มต้นสหกิจ/ฝึกงาน:")]
        public string date_start2 { get; set; }
        [DisplayName("วันที่สิ้นสุดสหกิจ/ฝึกงาน:")]
        public string date_end2 { get; set; }
        [DisplayName("จำนวนครั้งในการออกนิเทศ:")]
        [Required(ErrorMessage = "This field is required.")]
        public Nullable<int> to_go2 { get; set; }
        [DisplayName("วันที่เริ่มต้นสหกิจ/ฝึกงาน:")]
        public string date_start3 { get; set; }
        [DisplayName("วันที่สิ้นสุดสหกิจ/ฝึกงาน:")]
        public string date_end3 { get; set; }
        [DisplayName("จำนวนครั้งในการออกนิเทศ:")]
        [Required(ErrorMessage = "This field is required.")]
        public Nullable<int> to_go3 { get; set; }
    }
}