using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsApplication
{
    class SingleDataRecord
    {
        // Chromatogram
        public static string PRIMARY_ATTRIBUTE_NAME = "Chromatogram";
        public string ID = "";
        // RT [min]
        public static string SECOND_ATTRIBUTE_NAME = "RT [min]";
        public double time = 0;

        // other properties: Area, Int.Type, I, S/N, Max.m/z, FWHM [min], Area %
        public Dictionary<string, string> propertyMap = null;
        // key array, not contain the primary and second
        public ArrayList keyArray = null;

        public SingleDataRecord()
        {
            propertyMap = new Dictionary<string, string>();
            keyArray = new ArrayList();
        }

        public SingleDataRecord(string Chromatogram, double RTmin)
        {
            this.ID = Chromatogram;
            this.time = RTmin;
        }

        public bool AddProperty(string key, string value, string typeName)
        {
            if (propertyMap.ContainsKey(key + "-" + typeName))
            {
                return false;
            }
            else
            {
                propertyMap.Add(key + "-" + typeName, value);
                keyArray.Add(key + "-" + typeName);
                return true;
            }

        }
        public bool AddProperty(string key, string value)
        {
            if (propertyMap.ContainsKey(key))
            {
                return false;
            }
            else
            {
                propertyMap.Add(key, value);
                keyArray.Add(key);
                return true;
            }

        }
    }
}
