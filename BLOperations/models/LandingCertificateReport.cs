using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class LandingCertificateReport : table_Base
    {
        public string hbl_pkid { get; set; }
        public string hbl_no { get; set; }
        public string hbl_imp_name { get; set; }
        public string hbl_bl_no { get; set; }
        public string hbl_fcr_no { get; set; }
        public string hbl_date { get; set; }
        public string hbl_book_cntr { get; set; }
        public string mbl_pol_etd { get; set; }
        public string mbl_pol_name { get; set; }
        public Boolean hbl_no_checked { get; set; }
    }
}
