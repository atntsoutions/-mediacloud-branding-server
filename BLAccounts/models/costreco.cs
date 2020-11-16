using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Costreco : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }


        public string mbl_no { get; set; }
        public string mbl_bl_no { get; set; }
        public string mbl_book_no { get; set; }
        public string hbl_no { get; set; }
        public string hbl_bl_no { get; set; }
        public string agent_name { get; set; }
        public string hbl_type { get; set; }
        public string jvh_date { get; set; }
        public string jvh_type { get; set; }
        public string jvh_vrno { get; set; }
        public string acc_name { get; set; }
        public string jvh_cc_category { get; set; }
        public string ct_category { get; set; }
        public string jvh_remarks { get; set; }
        public string jvh_narration { get; set; }

        public string mstat{ get; set; }
        public string hstat { get; set; }

        public Nullable<decimal> acc_code { get; set; }
        public Nullable<decimal> jvh_cc_id { get; set; }
        public Nullable<decimal> jv_debit { get; set; }
        public Nullable<decimal> jv_credit { get; set; }

        public Nullable<decimal> jv_balance { get; set; }



        public Nullable<decimal> coldr0 { get; set; }
        public Nullable<decimal> colcr0 { get; set; }

        public Nullable<decimal> coldr1 { get; set; }
        public Nullable<decimal> colcr1 { get; set; }

        public Nullable<decimal> coldr2 { get; set; }
        public Nullable<decimal> colcr2 { get; set; }



    }
}
