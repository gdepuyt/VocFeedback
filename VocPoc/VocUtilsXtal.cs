using System.Collections.Generic;
using System.IO;
using System.Xml;
using VocPoC;
using static VocPoc.VocBrokerInformation;


namespace VocPoc
{
    /// <summary>
    /// This internal class provides utility methods for working with XML documents, specifically creating elements for a CRRequest schema.
    /// </summary>
    internal class XtalUtils
    {
        /// <summary>
        /// Creates a root element named "CRRequest" with the specified namespace URI in an XmlDocument.
        /// Sets default namespace attributes for xsi and xsd elements.
        /// </summary>
        /// <param name="doc">The XmlDocument to add the element to.</param>
        /// <param name="namespaceUri">The namespace URI for the element.</param>
        /// <returns>The created XmlElement representing the CRRequest root element.</returns>
        public static XmlElement CreateRootElement(XmlDocument doc, string namespaceUri)
        {
            XmlElement root = doc.CreateElement("CRRequest", namespaceUri);

            // Declare the default namespace for xsi and xsd elements
            root.SetAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            root.SetAttribute("xmlns:xsd", "http://www.w3.org/2001/XMLSchema");

            return root;
        }

        /// <summary>
        /// Creates a "RequestHeader" element containing child elements for CR identifier, type, owner information, external reference, priority, and communication count.
        /// </summary>
        /// <param name="doc">The XmlDocument to add the element to.</param>
        /// <param name="cridValue">The CR identifier value.</param>
        /// <param name="crTypeValue">The CR type value.</param>
        /// <param name="ownerEmailValue">The owner's email address.</param>
        /// <param name="ownerCostCenterValue">The owner's cost center.</param>
        /// <param name="externalReferenceValue">The external reference value.</param>
        /// <param name="priorityValue">The priority value.</param>
        /// <param name="comCountValue">The communication count value.</param>
        /// <returns>The created XmlElement representing the RequestHeader element.</returns>
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


        /// <summary>
        /// Creates a simple "Communications" element in the specified XmlDocument.
        /// </summary>
        /// <param name="doc">The XmlDocument to add the element to.</param>
        /// <returns>The created XmlElement representing the Communications element.</returns>
        public static XmlElement CreateCommunications(XmlDocument doc)
        {
            XmlElement Communications = doc.CreateElement("Communications");
            return Communications;
        }

        /// <summary>
        /// Creates a detailed "Communication" element in the specified XmlDocument.
        /// </summary>
        /// <param name="doc">The XmlDocument to add the element to.</param>
        /// <param name="communicationTypeValue">The type of communication (e.g., email, letter).</param>
        /// <param name="sequenceNumberValue">The sequence number for the communication.</param>
        /// <param name="addresseeLanguageValue">The language of the addressee.</param>
        /// <param name="addresseeRoleValue">The role of the addressee.</param>
        /// <param name="isColorValue">A boolean indicating whether the communication is in color (true) or black and white (false).</param>
        /// <param name="channelValue">The communication channel (e.g., email address, postal address).</param>
        /// <param name="uriValue">A Uniform Resource Identifier (URI) relevant to the communication.</param>
        /// <param name="fromValue">The sender information.</param>
        /// <param name="replyToValue">The reply-to address.</param>
        /// <param name="subjectValue">The subject of the communication.</param>
        /// <param name="attachmentFilePaths">An array of file paths for attachments</param>
        /// <param name="dynamicMailBD">A dictionary containing key-value pairs for dynamic elements in the BusinessData</param>
        /// <returns>The created XmlElement representing the Communication element.</returns>
        public static XmlElement CreateCommunication(XmlDocument doc, string communicationTypeValue, string sequenceNumberValue, string addresseeLanguageValue,
                                                     string addresseeRoleValue, string isColorValue, string channelValue, string uriValue,
                                                     string fromValue, string replyToValue, string subjectValue, string[] attachmentFilePaths, Dictionary<string, string> dynamicMailBD)
        {
            XmlElement communication = doc.CreateElement("Communication");



            XmlElement communicationType = doc.CreateElement("CommunicationType");
            communicationType.InnerText = communicationTypeValue;
            communication.AppendChild(communicationType);

            XmlElement sequenceNumber = doc.CreateElement("SequenceNumber");
            sequenceNumber.InnerText = sequenceNumberValue;
            communication.AppendChild(sequenceNumber);

            // TechnicalData
            XmlElement technicalData = doc.CreateElement("TechnicalData");
            communication.AppendChild(technicalData);

            XmlElement brInfo = doc.CreateElement("BRInfo");
            technicalData.AppendChild(brInfo);

            XmlElement addresseeLang = doc.CreateElement("AddresseeLanguage");
            addresseeLang.InnerText = addresseeLanguageValue;
            brInfo.AppendChild(addresseeLang);

            XmlElement addresseeRole = doc.CreateElement("AddresseeRole");
            addresseeRole.InnerText = addresseeRoleValue;
            brInfo.AppendChild(addresseeRole);

            XmlElement isColor = doc.CreateElement("IsColor");
            isColor.InnerText = isColorValue;
            brInfo.AppendChild(isColor);

            XmlElement locations = doc.CreateElement("Locations");
            technicalData.AppendChild(locations);

            XmlElement location = doc.CreateElement("Location");
            locations.AppendChild(location);

            XmlElement channel = doc.CreateElement("Channel");
            channel.InnerText = channelValue;
            location.AppendChild(channel);

            XmlElement uri = doc.CreateElement("URI");
            uri.InnerText = uriValue;
            location.AppendChild(uri);

            XmlElement from = doc.CreateElement("From");
            from.InnerText = fromValue;
            location.AppendChild(from);

            XmlElement replyTo = doc.CreateElement("ReplyTo");
            replyTo.InnerText = replyToValue;
            location.AppendChild(replyTo);

            XmlElement subject = doc.CreateElement("Subject");
            subject.InnerText = subjectValue;
            location.AppendChild(subject);

            XmlElement businessData = doc.CreateElement("BusinessData");
            communication.AppendChild(businessData);

            XmlElement structureData = doc.CreateElement("StructureData");
            businessData.AppendChild(structureData);

            XmlElement dynamicMailBusinessData = doc.CreateElement("DynamicMailBusinessData");
            structureData.AppendChild(dynamicMailBusinessData);

            // Add dynamic elements based on the dictionary (new section)
            if (dynamicMailBD != null && dynamicMailBD.Count > 0)
            {
                foreach (KeyValuePair<string, string> entry in dynamicMailBD)
                {
                    XmlElement element = doc.CreateElement(entry.Key);
                    element.InnerText = entry.Value;
                    dynamicMailBusinessData.AppendChild(element);
                }
            }

            // Attachments
            XmlElement attachments = doc.CreateElement("Attachments");
            communication.AppendChild(attachments);

            if (attachmentFilePaths != null && attachmentFilePaths.Length > 0)
            {
                int sequence = 1;
                foreach (string filePath in attachmentFilePaths)
                {
                    XmlElement attachment = doc.CreateElement("Attachment");
                    attachments.AppendChild(attachment);

                    XmlElement file = doc.CreateElement("File");
                    file.InnerText = filePath;
                    attachment.AppendChild(file);

                    XmlElement docType = doc.CreateElement("DocType");
                    // Set a default DocType, you can modify this based on your needs
                    docType.InnerText = "CO-BDAT";
                    attachment.AppendChild(docType);


                    sequence++;
                }
            }

            return communication;
        }

        /// <summary>
        /// Writes the content of an XmlDocument object to a specified file path in a formatted XML way.
        /// </summary>
        /// <param name="filePath">The full path to the file where the XML document will be saved.</param>
        /// <param name="doc">The XmlDocument object containing the XML data to be written.</param>
        /// <returns>An empty string upon successful execution (assuming no explicit return after writing).</returns>
        public static string WriteXmlToFile(string filePath, XmlDocument doc)
        {
            // Create XmlWriter settings with indentation
            var settings = new XmlWriterSettings
            {
                Indent = true,
                IndentChars = "\t",
                NewLineChars = "\n",
                CloseOutput = true
            };

            // Create XmlWriter for the specified file path
            using (XmlWriter writer = XmlWriter.Create(filePath, settings))
            {
                // Write the document content to the writer
                doc.WriteTo(writer);
                return writer.ToString();
            }
        }

        /// <summary>
        /// This class likely provides functionalities related to generating CR (Change Request) requests in an XML format.
        /// </summary>
        public class CRRequestGenerator
        {
            /// <summary>
            /// Initializes the generation of a CR request XML document. Retrieves configuration values, builds the XML content,
            /// and returns the generated XML string.
            /// </summary>
            /// <param name="FolderPath">The path to a folder (potentially used for output or configuration purposes).</param>
            /// <returns>The generated CR request XML document as a string.</returns>
            public static string CRRequestInit(string FolderPath)
            {
                // Retrieve configuration values
                string externalReference = VocConfigurationManagement.JobIdGenerator.JobId;
                string cridValue = VocConfigurationManagement.XtalConfig.VocXtal_crid;
                string crTypeValue = VocConfigurationManagement.XtalConfig.VocXtal_crType;
                string ownerEmailValue = VocConfigurationManagement.XtalConfig.VocXtal_ownerEmail;
                string ownerCostCenterValue = VocConfigurationManagement.XtalConfig.VocXtal_ownerCostCenter;
                string priorityValue = VocConfigurationManagement.XtalConfig.VocXtal_priority;
                string communicationTypeValue = VocConfigurationManagement.XtalConfig.VocXtal_communicationType;
                string addresseeRoleValue = VocConfigurationManagement.XtalConfig.VocXtal_addresseeRole;
                string isColorValue = VocConfigurationManagement.XtalConfig.VocXtal_isColor;
                string channelValue = VocConfigurationManagement.XtalConfig.VocXtal_channel;
                string fromValue = VocConfigurationManagement.XtalConfig.VocXtal_from;
                string replyToValue = VocConfigurationManagement.XtalConfig.VocXtal_replyTo;

                // Generate unique file path for output XML
                string fileOutputPath = Path.Combine(VocConfigurationManagement.FolderConfig.VocOutputXtalFileFolder,
                                                      VocFileManagement.AddUniqueStampToFileName("cr_voc.xml"));

                // Generate CR request XML content
                string generatedXml = GenerateCRRequestXml(
                    cridValue, crTypeValue, ownerEmailValue, ownerCostCenterValue, externalReference,
                    priorityValue, communicationTypeValue,
                    addresseeRoleValue, isColorValue, channelValue,
                    fromValue, replyToValue, fileOutputPath, FolderPath);

                return generatedXml;
            }

            /// <summary>
            /// Generates a CR (Change Request) request XML document based on provided configuration values.
            /// </summary>
            /// <param name="cridValue">The CR identifier.</param>
            /// <param name="crTypeValue">The CR type.</param>
            /// <param name="ownerEmailValue">The owner's email address.</param>
            /// <param name="ownerCostCenterValue">The owner's cost center.</param>
            /// <param name="externalReferenceValue">An external reference value.</param>
            /// <param name="priorityValue">The priority of the CR request.</param>
            /// <param name="communicationTypeValue">The communication type (e.g., email).</param>
            /// <param name="addresseeRoleValue">The role of the addressee.</param>
            /// <param name="isColorValue">A boolean indicating color preference (true) or black and white (false).</param>
            /// <param name="channelValue">The communication channel (e.g., email address).</param>
            /// <param name="fromValue">The sender information.</param>
            /// <param name="replyToValue">The reply-to address.</param>
            /// <param name="FileOutputPath">The full path for the output XML file.</param>
            /// <param name="folderPath">The path to a folder (potentially used for file processing).</param>
            /// <returns>The generated CR request XML document as a string.</returns>
            public static string GenerateCRRequestXml(string cridValue, string crTypeValue, string ownerEmailValue, string ownerCostCenterValue, string externalReferenceValue, string priorityValue,
                                                     string communicationTypeValue, string addresseeRoleValue, string isColorValue,
                                                     string channelValue, string fromValue, string replyToValue, string FileOutputPath, string folderPath)


            {
                XmlDocument doc = new XmlDocument();

                // Create root element with namespaces
                XmlElement root = CreateRootElement(doc, "http://tempuri.org"); // Replace with your actual namespace URI
                doc.AppendChild(root);

                // Add RequestHeader
                XmlElement requestHeader = CreateRequestHeader(doc, cridValue, crTypeValue, ownerEmailValue, ownerCostCenterValue, externalReferenceValue, priorityValue, "0");
                root.AppendChild(requestHeader);

                XmlElement originalRequestHeader = requestHeader;

                //Add Communications
                XmlElement communications = CreateCommunications(doc);
                root.AppendChild(communications);

                Dictionary<string, string> dynamicData = new Dictionary<string, string>();
                dynamicData.Add("MailSalutation", "Honoré"); // Replace with retrieved values
                                                             //dynamicData.Add("AnotherElement", "SomeValue"); // Add more elements as needed

                // Add Communications


                var fileInfoList = VocFileManagement.GetFilesInfo(folderPath);
                var sequenceComm = 1;


                foreach (var fileInfo in fileInfoList)
                {

                    BrokerDetails brokerDetails = VocBrokerInformation.GetBrokerDetails(fileInfo.FilenameWithoutExtension);

                    string subjectValue = VocTranslationManagement.FieldMappingProvider.GetTranslation("Email", "Title", brokerDetails.Language) + " " + brokerDetails.BrokerBCAB;
                    string addresseeLanguageValue = brokerDetails.Language;
                    string uriValue = brokerDetails.Email;
                    string sequenceNumberValue = sequenceComm.ToString();


                    string[] attachmentFilePaths = new string[1];

                    for (int i = 0; i < 1; i++)
                    {

                        attachmentFilePaths[i] = Path.Combine(VocConfigurationManagement.FolderConfig.VocOutputFileFolder, fileInfo.FileName);
                        //$@"\\zeus\a\softs\data\DUPL\VoC\Feedbacks\ToBrokers\{fileInfo.FilenameWithoutExtension}.xlsx";
                    }


                    //attachmentFilePaths = \\zeus\a\softs\data\DUPL\VoC\Feedbacks\ToBrokers\ fileInfo.FilenameWithoutExtension

                    XmlElement communication = CreateCommunication(doc, communicationTypeValue, sequenceNumberValue, addresseeLanguageValue, addresseeRoleValue, isColorValue,
                                                                channelValue, uriValue, fromValue, replyToValue, subjectValue, attachmentFilePaths, dynamicData);


                    communications.AppendChild(communication);

                    XmlElement sequenceNumberElement = (XmlElement)requestHeader.SelectSingleNode("//ComCount");

                    if (sequenceNumberElement != null)
                    {
                        sequenceNumberElement.InnerText = sequenceComm.ToString();
                    }


                    if (sequenceComm >= 10)
                    {
                        break; // Exit the loop after 10 iterations
                    }

                    sequenceComm++;

                }


                // Add Communications


                return WriteXmlToFile(FileOutputPath, doc);

            }

        }

    }
}