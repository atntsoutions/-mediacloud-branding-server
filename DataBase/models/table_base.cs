using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class table_Base
    {
        public string rec_company_code { get; set; }
        public string rec_branch_code { get; set; }
        public string rec_category { get; set; }


        public bool user_admin { get; set; }

        public string rec_created_by { get; set; }
        public string rec_created_date { get; set; }

        public string rec_edited_by { get; set; }
        public string rec_edited_date { get; set; }

        public Boolean rec_printed { get; set; }

        public Boolean rec_locked { get; set; }

        public Boolean rec_hidden { get; set; }


        public Boolean rec_deleted { get; set; }

        public string transfer_remarks { get; set; }

        public string approved_date { get; set; }
        public string approved_by { get; set; }
        public string approved_status { get; set; }
        public string approved_remarks { get; set; }

        public string rec_mode { get; set; }
        public string rec_aprvd_status { get; set; }
        public string rec_aprvd_remark { get; set; }
        public string rec_aprvd { get; set; }
        public string rec_aprvd_by { get; set; }

        public GlobalVariables  _globalvariables { get; set; }

    }


}