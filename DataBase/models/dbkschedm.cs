using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class DbkSchedm : table_Base
    {
        public string cntr_pkid { get; set; }
        public int cntr_slno { get; set; }
        public string cntr_no { get; set; }
        public decimal cntr_teu { get; set; }
        public string cntr_type_id { get; set; }
        public string cntr_type_code { get; set; }
        public string cntr_type_name { get; set; }
        public string cntr_csealno { get; set; }
        public string cntr_asealno { get; set; }
        public string cntr_booking_id { get; set; }
        public string cntr_oldbooking_id { get; set; }
        public string cntr_booking_no { get; set; }
        public string cntr_booking_name { get; set; }
        public string cntr_morh { get; set; }
        public string cntr_parent_id { get; set; }
        public string cntr_parent_type { get; set; }
        public decimal cntr_pcs { get; set; }
        public decimal cntr_grwt { get; set; }
        public decimal cntr_ntwt { get; set; }
        public decimal cntr_cbm { get; set; }
    }
}