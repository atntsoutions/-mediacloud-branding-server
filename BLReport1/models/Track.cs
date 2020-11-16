using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class TrackReport : table_Base
    {
        public int sl_no { get; set; }
        public string cntr_no { get; set; }
        public string hbl_bl_no { get; set; }
        public string vessel_name { get; set; }
        public string voyage { get; set; }
        public string pol_name { get; set; }
        public string pol_etd { get; set; }
        public string pol_etd_confirm { get; set; }
        public string pod_name { get; set; }
        public string pod_eta { get; set; }
        public string pod_eta_confirm { get; set; }
    }
}
