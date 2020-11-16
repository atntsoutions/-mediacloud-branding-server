using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{

    public class __Rebate : table_Base
    {
        public string row_type { get; set; }
        public string hbl_pkid { get; set; }
        public string inv_pkid { get; set; }
        public string inv_no { get; set; }
        public string hbl_no { get; set; }
        public string inv_source { get; set; }
        public string hbl_type { get; set; }
        public string mbl { get; set; }
        public string hbl { get; set; }
        public string created { get; set; }
        public string shipper_name { get; set; }
        public string inv_type { get; set; }
        public string acc_code { get; set; }
        public string acc_name { get; set; }
        public Nullable<decimal> inv_qty { get; set; }
        public Nullable<decimal> inv_rate { get; set; }
        public Nullable<decimal> inv_fototal { get; set; }
        public string inv_curr_code { get; set; }
        public Nullable<decimal> inv_exrate { get; set; }
        public Nullable<decimal> inv_total { get; set; }
        public Nullable<decimal> inv_rebate_amt { get; set; }
        public string inv_rebate_curr_code { get; set; }
        public Nullable<decimal> inv_rebate_exrate { get; set; }
        public Nullable<decimal> inv_rebate_amt_inr { get; set; }

        public string posted { get; set; }
        public Boolean selected { get; set; }

    }
}



