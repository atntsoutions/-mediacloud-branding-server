using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class CostPendingReport : table_Base
    {
        public string mbl_pkid { get; set; }
        public string mbl_hbl_no { get; set; }
        public string mbl_bl_no { get; set; }
        public string mbl_date { get; set; }
        public string mbl_sob_date { get; set; }
        public string mbl_folder_no { get; set; }
        public string mbl_book_cntr { get; set; }
        public string mbl_folder_sent_date { get; set; }
        public string mbl_agent_name { get; set; }
        public string mbl_nocosting { get; set; }
        public string cost_date{ get; set; }
        public string cost_refno { get; set; }
        public string cost_agent_name { get; set; }
        public string cost_folderno { get; set; }
        public string cost_jv_posted { get; set; }

    }
}
