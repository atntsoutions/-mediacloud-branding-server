using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{
    public class OsAgentReport : table_Base
    {
        public string rowtype { get; set; }
        public string rowcolor { get; set; }


        public string reccategory { get; set; }
        public string acc_code { get; set; }
        public string acc_name { get; set; }
        public string curr_code { get; set; }
        public Nullable<decimal> G1 { get; set; }
        public Nullable<decimal> G2 { get; set; }
        public Nullable<decimal> G3 { get; set; }
        public Nullable<decimal> G4 { get; set; }
        public Nullable<decimal> G5 { get; set; }
        public Nullable<decimal> balance { get; set; }
        public Nullable<decimal> advance { get; set; }

        public Nullable<decimal> G1_INR { get; set; }
        public Nullable<decimal> G2_INR { get; set; }
        public Nullable<decimal> G3_INR { get; set; }
        public Nullable<decimal> G4_INR { get; set; }
        public Nullable<decimal> G5_INR { get; set; }
        public Nullable<decimal> balance_inr { get; set; }
        public Nullable<decimal> advance_inr { get; set; }


        public Nullable<decimal> sea { get; set; }
        public Nullable<decimal> air { get; set; }
        public Nullable<decimal> oth { get; set; }
        public Nullable<decimal> adj { get; set; }
        public Nullable<decimal> bal { get; set; }


    }
}
