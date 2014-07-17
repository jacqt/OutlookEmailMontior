using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookEmailMonitor
{
    class Storage
    {
        public static void saveFile(String name, String data)
        {
            Properties.Settings.Default[name] = data;
            Properties.Settings.Default.Save();
        }

        public static String loadFile(String name)
        {
            object o = Properties.Settings.Default[name];
            if (o != null)
            {
                return o.ToString();
            }
            return null;
        }
    }
}
