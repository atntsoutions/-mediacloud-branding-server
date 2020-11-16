using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLHr
{
    public class HrReport : table_Base
    {
        public string row_type { get; set; }
        public string row_colour { get; set; }
        public string emp_no { get; set; }
        public string emp_name { get; set; }
        public string emp_pfno { get; set; }
        public Nullable<decimal> pf_base_salary { get; set; }
        public Nullable<decimal> pf_deduction { get; set; }
        public Nullable<decimal> emplyr_share { get; set; }
        public Nullable<decimal> pension { get; set; }
        public Nullable<decimal> vpf { get; set; }
        public Nullable<decimal> admin_chrg { get; set; }
        public Nullable<decimal> edli_chrg { get; set; }
        public Nullable<decimal> total_chrg { get; set; }
        public Nullable<decimal> eps_amt { get; set; }
        public string branch { get; set; }
        public string emp_esino { get; set; }
        public Nullable<decimal> sal_gross_earn { get; set; }
        public Nullable<decimal> emply_esi { get; set; }
        public Nullable<decimal> emplr_esi { get; set; }
        public Nullable<decimal> total { get; set; }
        public string edli_based_on { get; set; }
    }
}
