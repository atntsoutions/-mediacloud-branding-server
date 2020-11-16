using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataBase;
using DataBase_Oracle.Connections;


namespace BLEmail
{
    public class EmailService : BL_Base
    {
        public IDictionary<string, object> Process(Dictionary<string, object> SearchData)
        {
            string email_type ="";
            if (SearchData.ContainsKey("email_type"))
                email_type = SearchData["email_type"].ToString();

            if (email_type == "OS-ALL"|| email_type == "OS-DELHI" || email_type == "OS-SALESMAN-ALL")
            {//|| email_type == "OS-SALESMAN-ALL"
                using (OsService obj = new OsService())
                    return obj.ProcessOSReport(SearchData);
            }
            else if (email_type == "PAYSLIP-ALL")
            {
                using (PayslipService obj = new PayslipService())
                    return obj.Payslip_Mail(SearchData);
            }
            else 
            {
                Dictionary<string, object> RetData = new Dictionary<string, object>();
                RetData.Add("retvalue",false);
                RetData.Add("error", "Report Type Not Set (Email_Type)");
                return RetData;
            }
        }
          
    }

}
