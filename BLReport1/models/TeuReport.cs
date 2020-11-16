using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class TeuReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string cntr_pkid { get; set; }
        public string cntr_no { get; set; }
        public Nullable<decimal> cntr_teu { get; set; }
        public string cntr_csealno { get; set; }
        public string cntr_asealno { get; set; }
        public string cntr_type_code { get; set; }
        public string cntr_booking_no { get; set; }
        public string mbl_no { get; set; }
        public string mbl_book_no { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_shipment_type { get; set; }
        public string mbl_nature { get; set; }
        public string cntr_stuffed_at { get; set; }
        public string cntr_stuffed_on { get; set; }
        public Nullable<decimal> cntr_pcs { get; set; }
        public Nullable<decimal> cntr_ntwt { get; set; }
        public Nullable<decimal> cntr_grwt { get; set; }
        public Nullable<decimal> cntr_cbm { get; set; }
        public string cntr_clearing { get; set; }

        public string hbl_exp_name { get; set; }
        public string hbl_imp_name { get; set; }
        public string sman_name { get; set; }
        public string job_type { get; set; }

        public string hbl_agent_name { get; set; }
        public string hbl_carrier_name { get; set; }
        public string hbl_pol_name { get; set; }
        public string hbl_pod_name { get; set; }
        public string hbl_pofd_name { get; set; }
        public string branch { get; set; }
        public string hbl_nomination { get; set; }
    }
}
