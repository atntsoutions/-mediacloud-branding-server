using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataBase
{
    public class GlobalVariables
    {
        public string user_pkid { get; set; }
        public string user_code { get; set; }
        public string user_name { get; set; }
        public string user_email { get; set; }
        public string user_branch_id { get; set; }

        public string comp_pkid { get; set; }
        public string comp_code { get; set; }
        public string comp_name { get; set; }

        public string branch_pkid { get; set; }
        public string branch_code { get; set; }
        public string branch_name { get; set; }

        public string year_pkid { get; set; }
        public string year_code { get; set; }
        public string year_name { get; set; }
        public string year_prefix { get; set; }

        public string year_start_date { get; set; }
        public string year_end_date { get; set; }
        public string year_closed { get; set; }

        public string report_folder { get; set; }

        public string server_image_url { get; set; }
        public string server_image_path { get; set; }
        public string server_report_path { get; set; }

    }

}
