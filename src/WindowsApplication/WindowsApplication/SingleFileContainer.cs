using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsApplication
{
    class SingleFileContainer
    {
        private ArrayList recordList;

        public SingleFileContainer()
        {
            recordList = new ArrayList();
        }

        public void InsertRecord(SingleDataRecord newRecord)
        {
            this.recordList.Add(newRecord);
        }

        public void InsertFile(SingleFileContainer  newFile)
        {

        }
    }
}
