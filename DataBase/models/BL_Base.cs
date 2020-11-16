using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataBase
{
    public class BL_Base : IDisposable
    {
        public DataBase.Connections.DBConnection Con_Oracle = null;
        public string sql { get; set; }
        public void Dispose()
        {
            this.Con_Oracle = null;
        }
    }

   



}


