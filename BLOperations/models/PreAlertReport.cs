using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class PreAlertReport : table_Base
    {

        public string hbl_no { get; set; }
        public string hbl_agent_name { get; set; }
        public string hbl_shipper_name { get; set; }
        public string hbl_consignee_name { get; set; }
        public string hbl_consignee_add1 { get; set; }
        public string hbl_consignee_add2 { get; set; }
        public string hbl_consignee_add3 { get; set; }
        public string hbl_carrier_name { get; set; }
        public string hbl_bl_no { get; set; }
        public string hbl_pkg { get; set; }
        public string hbl_grwt { get; set; }
        public string hbl_ntwt { get; set; }
        public string hbl_chwt { get; set; }
        public string hbl_notify_name { get; set; }
        public string hbl_pod { get; set; }
        public string hbl_pofd { get; set; }
        public string hbl_commodity { get; set; }
        public string hawb_no { get; set; }
       
    }
}
