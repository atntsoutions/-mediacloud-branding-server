using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLReport1
{
    public class EgmReport : table_Base
    {
        public string cntr_no { get; set; }
        public string job_docno { get; set; }
        public string pol_code { get; set; }
        public string egno { get; set; }
        public string egdate { get; set; }
        public string opr_sbill_no { get; set; }
        public string opr_sbill_date { get; set; }

        public string job_cargo_nature { get; set; }
        public string pol { get; set; }
        public string pkg_unit { get; set; }
        public string pofd_code { get; set; }
        public string commodity { get; set; }
        public string exporter { get; set; }
        public string exp_add1 { get; set; }
        public string exp_add2 { get; set; }
        public string exp_add3 { get; set; }
        public string importer { get; set; }
        public string imp_add1 { get; set; }
        public string imp_add2 { get; set; }
        public string imp_add3 { get; set; }

        public Nullable<decimal> job_qty { get; set; }
        public Nullable<decimal> job_grwt { get; set; }
        public Nullable<decimal> job_pkg { get; set; }
        public Nullable<decimal> shut_out_qty { get; set; }


    }
    
}
