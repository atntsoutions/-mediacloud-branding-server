using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class OsRep : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }

        public string caption { get; set; }

        public string branch { get; set; }
        public string branch_code { get; set; }

        public string pkid { get; set; }

        public string smanid { get; set; }
        public string sman { get; set; }
        public string party { get; set; }
        public string invno { get; set; }
        public string invdate { get; set; }
        public string awbno { get; set; }

        public string jv_od_type { get; set; }
        public string jv_od_remarks { get; set; }

        public string jv_inv_category { get; set; }



        public Nullable<decimal> age1 { get; set; }
        public Nullable<decimal> age2 { get; set; }
        public Nullable<decimal> age3 { get; set; }
        public Nullable<decimal> age4 { get; set; }
        public Nullable<decimal> age5 { get; set; }
        public Nullable<decimal> age6 { get; set; }

        public Nullable<decimal> overdue { get; set; }
        public Nullable<decimal> balance { get; set; }
        public Nullable<decimal> advance { get; set; }
        public Nullable<decimal> legal { get; set; }

        public Nullable<decimal> oneyear { get; set; }



    }
}
