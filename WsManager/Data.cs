using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WsManager
{
    public static class Data
    {
        static Data()
        {
            InWb = new WbMan();
            OutWb = new WbMan();
        }

        public static WbMan InWb { get; set; }
        public static WbMan OutWb { get; set; }
    }
}
