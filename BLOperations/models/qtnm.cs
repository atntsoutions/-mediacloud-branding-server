using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLOperations
{
    public class Qtnm : table_Base
    {
        public string qtn_pkid { get; set; }
        public string qtn_type { get; set; }
        public string qtn_party_id { get; set; }
        public string qtn_remarks { get; set; }
        public decimal qtn_total { get; set; }

        public Qtnd Record { get; set; }

        public List<Qtnd> qtnList { get; set; }

    }

    public class Qtnd: table_Base
    {

        public string qtnd_pkid { get; set; }
        public string qtnd_parent_id { get; set; }

        public string qtnd_type { get; set; }


        public string qtnd_acc_main_code { get; set; }
        public string qtnd_acc_id { get; set; }
        public string qtnd_acc_code { get; set; }
        public string qtnd_acc_name { get; set; }

        public string qtnd_cntr_type_id { get; set; }
        public string qtnd_cntr_type_code { get; set; }

        public decimal qtnd_qty { get; set; }
        public decimal qtnd_rate { get; set; }

        public string qtnd_curr_id { get; set; }
        public string qtnd_curr_code { get; set; }
        public decimal qtnd_exrate { get; set; }
        public decimal qtnd_total { get; set; }

        public string qtnd_remarks { get; set; }
        public Int32 qtnd_ctr { get; set; }


    }


}
