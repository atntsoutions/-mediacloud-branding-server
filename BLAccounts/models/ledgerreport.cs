using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class LedgerReport : table_Base
    {
        public string pkid { get; set; }
        public string rowtype { get; set; }
        public string rowcolor { get; set; }

        public string level1 { get; set; }
        public string level2 { get; set; }
        public string grp_name { get; set; }
        public string acc_pkid { get; set; }
        public string acc_main_code { get; set; }
        public string acc_code { get; set; }
        public string acc_name { get; set; }
        public string sman_name { get; set; }
        public string cc_code { get; set; }
        public string cc_name { get; set; }
        public string cc_category { get; set; }
        public string cc_remarks { get; set; }
        public Nullable<decimal> cc_chwt { get; set; }
        public Nullable<decimal> cc_cbm { get; set; }

        public Int32 slno { get; set; }

        public string curr_code { get; set; }
        public string curr_id { get; set; }
        public Nullable<decimal> exrate { get; set; }


        public string jv_docno { get; set; }

        public string jv_type { get; set; }
        public string jv_vrno { get; set; }

        public string jv_year { get; set; }



        public string jv_date { get; set; }

        public string jv_org_invdt { get; set; }
        public string jv_ref_date { get; set; }

        public string jv_paid_date { get; set; }


        public string jv_narration { get; set; }
        public string jv_drcr { get; set; }

        public string jv_remarks { get; set; }

        public string jv_od_type { get; set; }
        public string jv_od_remarks { get; set; }

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


        public Nullable<decimal> age1 { get; set; }
        public Nullable<decimal> age2 { get; set; }
        public Nullable<decimal> age3 { get; set; }
        public Nullable<decimal> age4 { get; set; }
        public Nullable<decimal> age5 { get; set; }
       

        public Nullable<decimal> fdr { get; set; }
        public Nullable<decimal> fcr { get; set; }
        public Nullable<decimal> fbal { get; set; }
        public string fdrcr { get; set; }

        public string branch { get; set; }


        public Nullable<decimal> age6 { get; set; }
        public Nullable<decimal> oneyear { get; set; }
        public Nullable<decimal> overdue { get; set; }
        public Nullable<decimal> balance { get; set; }

        public string cust_pkid { get; set; }
        public string cust_name { get; set; }
        public string cust_code { get; set; }

        public string due_date { get; set; }


        public string bs_note_no { get; set; }
        public string bs_main_head { get; set; }
        public string bs_sub_head { get; set; }
        public string bs_sub_note { get; set; }

        public string cb_desc { get; set; }
        public Nullable<decimal> cb_dr { get; set; }
        public Nullable<decimal> cb_cr { get; set; }



        public Nullable<decimal> apr { get; set; }
        public Nullable<decimal> may { get; set; }
        public Nullable<decimal> jun { get; set; }
        public Nullable<decimal> jul { get; set; }
        public Nullable<decimal> aug { get; set; }
        public Nullable<decimal> sep { get; set; }

        public Nullable<decimal> oct { get; set; }
        public Nullable<decimal> nov { get; set; }
        public Nullable<decimal> dec { get; set; }

        public Nullable<decimal> jan { get; set; }
        public Nullable<decimal> feb { get; set; }
        public Nullable<decimal> mar { get; set; }


        public Nullable<decimal> row_debit { get; set; }
        public Nullable<decimal> row_credit { get; set; }


    }
}
