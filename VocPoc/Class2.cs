using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace MyTestEpPLus
{
    internal class XtalUtils
    {
        public static XmlElement CreateRequestHeader(XmlDocument doc, string cridValue, string crTypeValue, string ownerEmailValue, string ownerCostCenterValue, string externalReferenceValue, string priorityValue, string comCountValue)
        {
            XmlElement requestHeader = doc.CreateElement("RequestHeader");

            XmlElement crid = doc.CreateElement("CRID");
            crid.InnerText = cridValue;
            requestHeader.AppendChild(crid);

            XmlElement crType = doc.CreateElement("CRType");
            crType.InnerText = crTypeValue;
            requestHeader.AppendChild(crType);

            XmlElement owner = doc.CreateElement("Owner");
            requestHeader.AppendChild(owner);

            XmlElement ownerEmail = doc.CreateElement("EMailAddress");
            ownerEmail.InnerText = ownerEmailValue;
            owner.AppendChild(ownerEmail);

            XmlElement ownerCostCenter = doc.CreateElement("CostCenter");
            ownerCostCenter.InnerText = ownerCostCenterValue;
            owner.AppendChild(ownerCostCenter);

            XmlElement externalRef = doc.CreateElement("ExternalReference");
            externalRef.InnerText = externalReferenceValue;
            requestHeader.AppendChild(externalRef);

            XmlElement priority = doc.CreateElement("Priority");
            priority.InnerText = priorityValue;
            requestHeader.AppendChild(priority);

            XmlElement comCount = doc.CreateElement("ComCount");
            comCount.InnerText = comCountValue;
            requestHeader.AppendChild(comCount);

            return requestHeader;
        }
    }
}
