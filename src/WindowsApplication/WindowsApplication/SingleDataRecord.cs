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
        private string ID;
        // RT [min]
        private double time;

        // other properties: Area, Int.Type, I, S/N, Max.m/z, FWHM [min], Area %
        private Dictionary<string, string> propertyMap;

        public SingleDataRecord(string Chromatogram, double RTmin)
        {
            this.ID = Chromatogram;
            this.time = RTmin;
        }

        public bool addProperty(string key, string value, string typeName)
        {
            if (propertyMap.ContainsKey(typeName + "-" + key))
            {
                return false;
            }
            else
            {
                propertyMap.Add(typeName + "-" + key, value);
                return true;
            }

        }
    }
}
