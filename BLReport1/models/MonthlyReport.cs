using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class MonthlyReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string sino { get; set; }
        public string folder_no { get; set; }
        public string folder_sent { get; set; }
        public string mbl_no { get; set; }
        public string mbl_date { get; set; }
        public string mbl_status { get; set; }
        public string hbl_no { get; set; }
        public string hbl_date { get; set; }
        public string hbl_status { get; set; }
        public string shipper_name { get; set; }
        public string consignee_name { get; set; }
        public string agent_name { get; set; }
        public string hbl_nomination { get; set; }
        public string carrier_name { get; set; }
        public string pol_name { get; set; }
        public string pod_name { get; set; }
        public string pofd_name { get; set; }
        public string pol_etd { get; set; }
        public string sman_name { get; set; }
        public Nullable<decimal> hbl_grwt { get; set; }
        public Nullable<decimal> hbl_chwt { get; set; }
        public Nullable<decimal> mbl_grwt { get; set; }
        public Nullable<decimal> mbl_chwt { get; set; }
        public string netnet { get; set; }
        public Nullable<decimal> publish_rate { get; set; }
        public Nullable<decimal> informed_rate { get; set; }
        public Nullable<decimal> sell_informed { get; set; }
        public Nullable<decimal> rebate { get; set; }
        public Nullable<decimal> exworks { get; set; }
        public string commodty_name { get; set; }

        public string hbl_type { get; set; }
        public Nullable<decimal> hbl_book_cntr_teu { get; set; }
        public Nullable<decimal> hbl_cbm { get; set; }

        public string hbl_nature { get; set; }
        public string mbl_nature { get; set; }
        public string hbl_terms { get; set; }
        public string mbl_terms { get; set; }
        public string shipment_type { get; set; }
        public string hbl_book_cntr { get; set; }
        public string hbl_job_nos { get; set; }
        public Nullable<decimal> hbl_ntwt { get; set; }

        public string hbl_ar_invnos { get; set; }
        public Nullable<decimal> hbl_ar_invamt { get; set; }
        public Nullable<decimal> hbl_ar_gstamt { get; set; }

        public string hbl_pkid { get; set; }
        public string sman_id { get; set; }
        public bool displayed { get; set; }


        public string branch { get; set; }
        public string agent_created_date { get; set; }

        public string created_date { get; set; }


    }
}
