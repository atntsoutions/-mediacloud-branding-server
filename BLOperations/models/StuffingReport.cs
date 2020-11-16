using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class StuffingReport : table_Base
    {
        public string job_exp_name { get; set; }
        public string job_imp_name { get; set; }
        public string job_pofd_name { get; set; }
        public string opr_sbill_no { get; set; }
        public string opr_sbill_date { get; set; }
        public Nullable<decimal> pack_pkg { get; set; }
        public Nullable<decimal> pack_cbm { get; set; }
        public Nullable<decimal> pack_grwt { get; set; }
        public string pack_cntr_no { get; set; }
        public string pack_cntr_csealno { get; set; }
        public string pack_cntr_asealno { get; set; }
        public string mbl_vessel_name { get; set; }
        public string mbl_vessel_voyage { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_pol_etd_confirm { get; set; }
        public string mbl_pod_eta { get; set; }
        public string itm_order_no { get; set; }
        public string itm_uneco { get; set; }
        public string itm_styleno { get; set; }
        
        
        

    }
}
