using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class BkTeuReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string cntr_pkid { get; set; }
        public string cntr_no { get; set; }             
        public string cntr_type_code { get; set; }
        public string cntr_booking_no { get; set; }
        public string mbl_no { get; set; }
        public string mbl_book_no { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_shipment_type { get; set; }
        public string mbl_nature { get; set; }
        public Nullable<decimal> mbl_book_cntr_mcbm { get; set; }
        public Nullable<decimal> mbl_book_cntr_m20 { get; set; }
        public Nullable<decimal> mbl_book_cntr_m40 { get; set; }
        public Nullable<decimal> mbl_book_cntr_mteu { get; set; }

        public string mbl_exp_name { get; set; }
        public string mbl_imp_name { get; set; }
        public string mbl_carrier_name { get; set; }

        public string cntr_clearing { get; set; }
        public string hbl_pol { get; set; }
        public string hbl_pod { get; set; }
        public string hbl_pofd { get; set; }
        public string hbl_agent { get; set; }
        public string hbl_nomination { get; set; }
        
        public string branch { get; set; }

    }
}
