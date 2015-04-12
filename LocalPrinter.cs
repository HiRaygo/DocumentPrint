using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace DocumentPrint
{
    internal class LocalPrinter
    {
        public LocalPrinter(string deviceID)
        {
            this.DeviceID = deviceID;
        }

        public string DeviceID { get; set; }
        public string DevicePortAddress { get; set; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            byte propertyCount = 0;
            foreach (PropertyInfo property in this.GetType().GetProperties())
            {
                if (property.CanRead)
                {
                    if (propertyCount > 0)
                    {
                        sb.Append(", ");
                    }
                    sb.Append(property.Name);
                    sb.Append(" = ");
                    sb.Append(property.GetValue(this, null));

                    propertyCount++;
                }
            }
            return sb.ToString();
        }
    }
}
