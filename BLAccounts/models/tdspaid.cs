using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class TdsPaid : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }

        public string jvh_date { get; set; }
        public string jvh_vrno { get; set; }
        public string jvh_type { get; set; }
       
        public string party_code { get; set; }
        public string party_name { get; set; }
        public string tan { get; set; }
        public string tan_name { get; set; }
        public string sman_name { get; set; }
        public string jvh_narration { get; set; }

        public Nullable<decimal> jv_gross_bill_amt { get; set; }
        public Nullable<decimal> jv_debit { get; set; }
        public Nullable<decimal> jv_credit { get; set; }

        

    }
}
