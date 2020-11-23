using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class User : table_Base
    {
        public string user_pkid { get; set; }
        public string user_code { get; set; }
        public string user_name { get; set; }

        public string user_parent_id { get; set; }
        public string user_parent_name { get; set; }

        public string user_password { get; set; }
        public string user_email { get; set; }
        public string user_email_pwd { get; set; }

        public string user_company_id { get; set; }
        public string user_company_code { get; set; }
        public string user_company_name { get; set; }

        public string user_branch_id { get; set; }
        public string user_branch_name { get; set; }
        public int user_rights_total { get; set; }

        public string user_sman_id { get; set; }
        public string user_sman_code { get; set; }
        public string user_sman_name { get; set; }

        public string user_role_id { get; set; }
        public string user_role_name { get; set; }
        public string user_role_rights_id { get; set; }


        public string user_vendor_id { get; set; }
        public string user_vendor_name { get; set; }

        public string user_region_id { get; set; }
        public string user_region_name { get; set; }


        public string user_local_server { get; set; }

        public string user_token_id { get; set; }
        public string user_ipaddress { get; set; }

        public Boolean user_branch_user { get; set; }

        public Boolean user_islocked { get; set; }

        public List<Userd> recorddet { get; set; }

    }

}