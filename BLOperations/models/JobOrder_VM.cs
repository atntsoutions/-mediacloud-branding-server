using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLOperations
{
    public class JobOrder_VM 
    {
        public GlobalVariables globalVariables { get; set; }
        public List<Joborderm> JobOrder { get; set; }
        public string ord_exp_id { get; set; }
        public string ord_imp_id { get; set; }
        public string ord_agent_id { get; set; }
        public string ord_parent_id { get; set; }
        public string ord_source { get; set; }
    }

   

}
