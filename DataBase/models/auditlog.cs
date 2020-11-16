using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Auditlog : table_Base
    {
        public string audit_date { get; set; }
        public string audit_module { get; set; }
        public string audit_type { get; set; }
        public string audit_action { get; set; }
        public string audit_comp_code { get; set; }
        public string audit_branch_code { get; set; }
        public string audit_user_code { get; set; }
        public string audit_pkey { get; set; }
        public string audit_refno { get; set; }
        public string audit_computer { get; set; }
        public decimal audit_old_amt { get; set; }
        public decimal audit_new_amt { get; set; }
        public string audit_old_remarks { get; set; }
        public string audit_remarks { get; set; }

    }


}

