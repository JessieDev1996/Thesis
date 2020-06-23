using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Project_Test1.Models
{
    public class SupDoc
    {
        public string title_en { get; set; }
        public string title_th { get; set; }
        public string title_detail { get; set; }
        public string job_position { get; set; }
        public string job_description { get; set; }
        public string apartment { get; set; }
        public string room { get; set; }
        public string receive_doc { get; set; }
        public string name { get; set; }
        public string position { get; set; }
        public string department { get; set; }
        public string phone { get; set; }
        public string fax { get; set; }
        public string email { get; set; }

        public int prototype_id { get; set; }
        public string number { get; set; }
        public string building { get; set; }
        public string floor { get; set; }
        public string lane { get; set; }
        public string road { get; set; }
    }
}