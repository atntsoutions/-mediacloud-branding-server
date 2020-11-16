using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class VesselLoadingReport : table_Base
    {
        public string mbl_agent_name { get; set; }
        public string mbl_vessel_name { get; set; }
        public string mbl_vessel_voyage { get; set; }
        public string book_cntr_no { get; set; }
        public Nullable<decimal> book_cntr_grwt { get; set; }
        public string book_cntr_grwt_unit { get; set; }
        public string book_cntr_type { get; set; }
        public string mbl_commodity { get; set; }
        public string mbl_pofd_name { get; set; }
        public string book_cntr_stuffed_at { get; set; }
        public string book_cntr_stuffed_on { get; set; }
        public string mbl_book_no { get; set; }
        
    }
}
