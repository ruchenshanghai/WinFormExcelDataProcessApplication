using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsApplication
{
    class SingleDataRecord
    {
        // Chromatogram
        public string ID = "";
        // RT [min]
        public double time = 0;

        // other properties: Area, Int.Type, I, S/N, Max.m/z, FWHM [min], Area %
        public Dictionary<string, string> propertyMap = null;

        public SingleDataRecord()
        {
            propertyMap = new Dictionary<string, string>();
        }

        //public SingleDataRecord(string Chromatogram, double RTmin)
        //{
        //    this.ID = Chromatogram;
        //    this.time = RTmin;
        //}

        public bool AddProperty(string key, string value, string typeName)
        {
            if (propertyMap.ContainsKey(key + "-" + typeName))
            {
                return false;
            }
            else
            {
                propertyMap.Add(key + "-" + typeName, value);
                return true;
            }

        }
    }
}
