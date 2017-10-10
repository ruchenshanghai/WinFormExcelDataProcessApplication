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

        public SingleFileContainer MergeFileContainer(SingleFileContainer  newContainer)
        {
            SingleFileContainer resultContainer = new SingleFileContainer();
            for (int outerIndex = 0; outerIndex < this.recordList.Count; outerIndex++)
            {
                SingleDataRecord tempOuterRecord = (SingleDataRecord)recordList[outerIndex];
                for (int innerIndex = 0; innerIndex < newContainer.recordList.Count; innerIndex++)
                {
                    // compare the ID first, then compare the time

                }
            }


            return null;
        }
    }
}
