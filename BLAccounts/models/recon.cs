using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class Recon : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }

        public string level1 { get; set; }
        public string level2 { get; set; }
        public string recon_grp_name { get; set; }
        public string recon_acc_pkid { get; set; }
        public string recon_acc_main_code { get; set; }
        public string recon_acc_code { get; set; }
        public string recon_acc_name { get; set; }
        public string recon_cc_code { get; set; }
        public string recon_cc_name { get; set; }
        public string recon_cc_category { get; set; }
        public string recon_cc_remarks { get; set; }

        
            
        public Int32 slno { get; set; }

        public string recon_jv_vrno { get; set; }
        public string recon_jv_type { get; set; }

        public string recon_jv_year { get; set; }
        public string recon_jv_docno { get; set; }
        public string recon_jv_date { get; set; }
        public string recon_jv_narration { get; set; }
        public string recon_jv_drcr { get; set; }
        public string recon_chqno { get; set; }
        public string recon_due_date { get; set; }


        public string jvh_not_over_chq { get; set; }

        public string jv_pkid { get; set; }

        public string recon_date { get; set; }
        public string recon_display_date { get; set; }


        public Nullable<decimal> recon_amount { get; set; }
        public string recon_amount_type { get; set; }

        public string recon_paid_to { get; set; }
        public string recon_bank { get; set; }
        public string recon_type { get; set; }
        public Nullable<decimal> opbal { get; set; }
        public Nullable<decimal> bal { get; set; }


        public Nullable<decimal> debit { get; set; }
        public Nullable<decimal> credit { get; set; }

        public Nullable<decimal> drbal { get; set; }
        public Nullable<decimal> crbal { get; set; }

        public Nullable<decimal> advance { get; set; }


        public Nullable<decimal> crdays { get; set; }
        public Nullable<decimal> crlimit { get; set; }
        public Nullable<decimal> osdays { get; set; }
        public Nullable<decimal> overduedays { get; set; }

    }
}
