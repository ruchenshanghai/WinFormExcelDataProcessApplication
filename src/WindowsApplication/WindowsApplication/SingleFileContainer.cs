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
        public ArrayList recordList;
        public ArrayList keyArray = null;
        public static string DEFAULT_VALUE = "XXX";

        public SingleFileContainer()
        {
            recordList = new ArrayList();
        }

        public void InsertRecord(SingleDataRecord newRecord)
        {
            this.recordList.Add(newRecord);
            if (keyArray == null)
            {
                keyArray = newRecord.keyArray;
            }
        }

        public SingleFileContainer MergeFileContainer(SingleFileContainer  newContainer, double rangeParam)
        {
            double rangeDelta = rangeParam;
            SingleFileContainer resultContainer = new SingleFileContainer();
            for (int outerIndex = 0; outerIndex < this.recordList.Count; outerIndex++)
            {
                SingleDataRecord tempOuterRecord = (SingleDataRecord)recordList[outerIndex];
                if (tempOuterRecord.isMerged)
                {
                    // has merged
                    continue;
                }
                //bool isMatch = false;
                for (int innerIndex = 0; innerIndex < newContainer.recordList.Count; innerIndex++)
                {

                    // compare the ID first, then compare the time
                    SingleDataRecord tempInnerRecord = (SingleDataRecord)newContainer.recordList[innerIndex];
                    if (tempInnerRecord.isMerged)
                    {
                        // has merged
                        continue;
                    }
                    if ((tempOuterRecord.ID == tempInnerRecord.ID) && (Math.Abs(tempOuterRecord.time - tempInnerRecord.time) <= rangeDelta))
                    {
                        double averageTime = (tempOuterRecord.time + tempInnerRecord.time) / 2;
                        SingleDataRecord tempResultRecord = new SingleDataRecord(tempOuterRecord.ID, averageTime);
                        // copy other key-value from inner and outter
                        for (int keyIndex = 0; keyIndex < tempOuterRecord.keyArray.Count; keyIndex++)
                        {
                            string tempKey = (string)tempOuterRecord.keyArray[keyIndex];
                            string tempValue = tempOuterRecord.propertyMap[tempKey];
                            if (!tempResultRecord.AddProperty(tempKey, tempValue))
                            {
                                // add property failed
                                return null;
                            }
                        }
                        for (int keyIndex = 0; keyIndex < tempInnerRecord.keyArray.Count; keyIndex++)
                        {
                            string tempKey = (string)tempInnerRecord.keyArray[keyIndex];
                            string tempValue = tempInnerRecord.propertyMap[tempKey];
                            if (!tempResultRecord.AddProperty(tempKey, tempValue))
                            {
                                // add property failed
                                return null;
                            }
                        }

                        // add to result, remove previous data
                        resultContainer.InsertRecord(tempResultRecord);
                        ((SingleDataRecord)this.recordList[outerIndex]).isMerged = true;
                        ((SingleDataRecord)newContainer.recordList[innerIndex]).isMerged = true;
                        break;
                    }
                }
            }
            // deal with the rest this data
            for (int recordIndex = 0; recordIndex < this.recordList.Count; recordIndex++)
            {
                SingleDataRecord rawDataRecord = (SingleDataRecord)this.recordList[recordIndex];
                if (rawDataRecord.isMerged)
                {
                    // has merged
                    continue;
                }
                SingleDataRecord tempResultRecord = new SingleDataRecord(rawDataRecord.ID, rawDataRecord.time);
                for (int newIndex = 0; newIndex < rawDataRecord.keyArray.Count; newIndex++)
                {
                    string tempKey = (string)rawDataRecord.keyArray[newIndex];
                    string tempValue = rawDataRecord.propertyMap[tempKey];
                    if (!tempResultRecord.AddProperty(tempKey, tempValue))
                    {
                        // add property failed
                        return null;
                    }
                }
                for (int newIndex = 0; newIndex < newContainer.keyArray.Count; newIndex++)
                {
                    string tempKey = (string)newContainer.keyArray[newIndex];
                    string tempValue = DEFAULT_VALUE;
                    if (!tempResultRecord.AddProperty(tempKey, tempValue))
                    {
                        // add property failed
                        return null;
                    }
                }
                resultContainer.InsertRecord(tempResultRecord);
                rawDataRecord.isMerged = true;
            }
            // deal with the rest new data
            for (int recordIndex = 0; recordIndex < newContainer.recordList.Count; recordIndex++)
            {
                SingleDataRecord rawDataRecord = (SingleDataRecord)newContainer.recordList[recordIndex];
                if (rawDataRecord.isMerged)
                {
                    // has merged
                    continue;
                }
                SingleDataRecord tempResultRecord = new SingleDataRecord(rawDataRecord.ID, rawDataRecord.time);
                for (int newIndex = 0; newIndex < rawDataRecord.keyArray.Count; newIndex++)
                {
                    string tempKey = (string)rawDataRecord.keyArray[newIndex];
                    string tempValue = rawDataRecord.propertyMap[tempKey];
                    if (!tempResultRecord.AddProperty(tempKey, tempValue))
                    {
                        // add property failed
                        return null;
                    }
                }
                for (int newIndex = 0; newIndex < this.keyArray.Count; newIndex++)
                {
                    string tempKey = (string)this.keyArray[newIndex];
                    string tempValue = DEFAULT_VALUE;
                    if (!tempResultRecord.AddProperty(tempKey, tempValue))
                    {
                        // add property failed
                        return null;
                    }
                }
                resultContainer.InsertRecord(tempResultRecord);
                rawDataRecord.isMerged = true;
            }


            for (int i = 0; i < this.recordList.Count; i++)
            {
                if (!((SingleDataRecord)this.recordList[i]).isMerged)
                {
                    Console.WriteLine("Error");
                }
            }
            for (int i = 0; i < newContainer.recordList.Count; i++)
            {
                if (!((SingleDataRecord)newContainer.recordList[i]).isMerged)
                {
                    Console.WriteLine("Error");
                }
            }

            return resultContainer;
        }
    }
}
