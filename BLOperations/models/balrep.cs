
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace BLOperations
{
    public class BalRep : table_Base
    {
        public string brcode { get; set; }
        public decimal creditamt { get; set; }
        public decimal overdueamt { get; set; }
        public decimal overduedays { get; set; }
    }

}
