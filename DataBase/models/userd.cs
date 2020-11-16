using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{


    public class Userd : table_Base
    {
        public string user_id { get; set; }
        public string user_branch_id { get; set; }
        public string user_company_name { get; set; }
        public string user_branch_name { get; set; }
        public Boolean user_selected { get; set; }
        public Boolean user_default_branch_id { get; set; }
    }

}