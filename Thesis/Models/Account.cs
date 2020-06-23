using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace Project_Test1.Models
{
    public class Account
    {
        public string user_id { get; set; }
        public string password { get; set; }
        public string position { get; set; }
        public string major { get; set; }
        public string date_time_in { get; set; }
        public string date_time_update { get; set; }

        //for alert
        public string loginErrMsg { get; set; }
    }
}