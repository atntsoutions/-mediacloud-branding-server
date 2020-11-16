using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class ArrivalNotice : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string mbl_pkid { get; set; }
        public string mbl_slno { get; set; }
        public string mbl_no { get; set; }
        public string mbl_book_no { get; set; }
        public string mbl_book_date { get; set; }
        public string hbl_blno { get; set; }
        public string mbl_pol_code { get; set; }
        public string mbl_pol_name { get; set; }
        public string mbl_pod_code { get; set; }
        public string mbl_pod_name { get; set; }
        public string mbl_carrier_name { get; set; }
        public string hbl_pkid { get; set; }
        public string hbl_exp_code { get; set; }
        public string hbl_exp_name { get; set; }
        public string hbl_exp_br_no { get; set; }
        public string hbl_exp_addr1 { get; set; }
        public string hbl_exp_addr2 { get; set; }
        public string hbl_exp_addr3 { get; set; }
        public string hbl_imp_code { get; set; }
        public string hbl_imp_name { get; set; }
        public string hbl_imp_br_no { get; set; }
        public string hbl_imp_addr1 { get; set; }
        public string hbl_imp_addr2 { get; set; }
        public string hbl_imp_addr3 { get; set; }
        public string hbl_terms { get; set; }
        public string mbl_vessel_code { get; set; }
        public string mbl_vessel_name { get; set; }
        public string mbl_vessel_no { get; set; }
        public string hbl_commodity { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_pod_eta { get; set; }
        public string hbl_cntrs { get; set; }
        public string hbl_ar_notice { get; set; }
        public int mbl_eta_days { get; set; }
        public Nullable<decimal> hbl_packages { get; set; }
        public Nullable<decimal> hbl_grweight { get; set; }
        public Nullable<decimal> hbl_volume { get; set; }


    }
}
