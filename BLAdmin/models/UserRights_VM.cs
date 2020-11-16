using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DataBase;

namespace DataBase
{

    public class UserRights_VM
    {
        public GlobalVariables globalVariables { get; set; }
        public List<UserRights> userRights { get; set; }
        public string copyto_user_id { get; set; }
        public string copyto_branch_id { get; set; }
    }




}