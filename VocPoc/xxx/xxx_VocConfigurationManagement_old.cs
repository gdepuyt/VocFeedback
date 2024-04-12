using log4net;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VocPoc
{
    public class xxx_VocConfigurationManagement_old
    {
        private static readonly ILog _log = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public static void LogAllSettings()
        {
            var appSettings = ConfigurationManager.AppSettings;
            foreach (var key in appSettings.AllKeys)
            {
                string value = appSettings[key];
                _log.Info($"App setting: {key} = {value}"); // Replace with your logging method
            }
        }
        public static class FolderConfig
        {
            private static string _vocRootFolder;
            private static string _vocFeedbackFileFolder;
            private static string _vocWorkingFileFolder;
            private static string _vocOutputFileFolder;
            private static string _vocOutputXtalFileFolder;
            private static string _vocResponseFilenamePattern;
            private static string _vocBrokerFilenamePattern;
            private static string _vocTemplateFileFolder;

            static FolderConfig()
            {
                try
                {
                    _vocRootFolder = ConfigurationManager.AppSettings["VocRootFolder"];
                    _vocFeedbackFileFolder = ConfigurationManager.AppSettings["VocFeedbackFileFolder"];
                    _vocWorkingFileFolder = ConfigurationManager.AppSettings["VocWorkingFileFolder"];
                    _vocOutputFileFolder = ConfigurationManager.AppSettings["VocOutputFileFolder"];
                    _vocOutputXtalFileFolder = ConfigurationManager.AppSettings["VocOutputXtalFileFolder"];
                    _vocResponseFilenamePattern = ConfigurationManager.AppSettings["VocResponseFilenamePattern"];
                    _vocBrokerFilenamePattern = ConfigurationManager.AppSettings["VocBrokerFilenamePattern"];
                    _vocTemplateFileFolder = ConfigurationManager.AppSettings["VocTemplateFileFolder"];

                    if (string.IsNullOrEmpty(_vocRootFolder) ||
                        string.IsNullOrEmpty(_vocFeedbackFileFolder) ||
                        string.IsNullOrEmpty(_vocWorkingFileFolder) ||
                        string.IsNullOrEmpty(_vocOutputFileFolder) ||
                        string.IsNullOrEmpty(_vocOutputXtalFileFolder) ||
                        string.IsNullOrEmpty(_vocResponseFilenamePattern) ||
                        string.IsNullOrEmpty(_vocBrokerFilenamePattern) ||
                        string.IsNullOrEmpty(_vocTemplateFileFolder))
                    {
                        throw new ConfigurationErrorsException("Missing folder paths in app.config");
                    }

                    // Test folder existence
                    ValidateFolderExistence(_vocRootFolder);
                    ValidateFolderExistence(_vocFeedbackFileFolder);
                    ValidateFolderExistence(_vocWorkingFileFolder);
                    ValidateFolderExistence(_vocOutputFileFolder);
                    ValidateFolderExistence(_vocOutputXtalFileFolder);
                    ValidateFolderExistence(_vocTemplateFileFolder);

                }
                catch (ConfigurationErrorsException ex)
                {
                    _log.Error($"Error reading folder paths from app.config: {ex.Message}");
                    throw; // Or re-throw the exception for handling by the caller
                }
                catch (DirectoryNotFoundException ex)
                {
                    _log.Error($"Folder not found: {ex.Message}");
                    throw; // Or handle the missing folder error gracefully
                }
            }

            private static void ValidateFolderExistence(string folderPath)
            {
                if (!Directory.Exists(folderPath))
                {
                    throw new DirectoryNotFoundException($"Folder '{folderPath}' does not exist.");
                }
            }


            public static string[] FolderStringToArray(string commaSeparatedFilenames)
            {
                if (string.IsNullOrEmpty(commaSeparatedFilenames))
                {
                    return Array.Empty<string>(); // Handle empty string or null value
                }

                return commaSeparatedFilenames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            }


            public static string VocRootFolder => _vocRootFolder;
            public static string VocFeedbackFileFolder => _vocFeedbackFileFolder;
            public static string VocWorkingFileFolder => _vocWorkingFileFolder;
            public static string VocOutputFileFolder => _vocOutputFileFolder;
            public static string VocOutputXtalFileFolder => _vocOutputXtalFileFolder;
            public static string VocResponseFilenamePattern => _vocResponseFilenamePattern;
            public static string VocBrokerFilenamePattern => _vocBrokerFilenamePattern;

            public static string VocTemplateFileFolder => _vocTemplateFileFolder;
        }
        public static class XtalConfig
        {
            private static string _vocXtal_crid;
            private static string _vocXtal_crType;
            private static string _vocXtal_ownerEmail;
            private static string _vocXtal_ownerCostCenter;
            private static string _vocXtal_priority;
            private static string _vocXtal_communicationType;
            private static string _vocXtal_addresseeRole;
            private static string _vocXtal_isColor;
            private static string _vocXtal_channel;
            private static string _vocXtal_from;
            private static string _vocXtal_replyTo;


            static XtalConfig()
            {
                try
                {

                    _vocXtal_crid = ConfigurationManager.AppSettings["Voc_Xtal_crid"];
                    _vocXtal_crType = ConfigurationManager.AppSettings["Voc_Xtal_crType"];
                    _vocXtal_ownerEmail = ConfigurationManager.AppSettings["Voc_Xtal_ownerEmail"];
                    _vocXtal_ownerCostCenter = ConfigurationManager.AppSettings["Voc_Xtal_ownerCostCenter"];
                    _vocXtal_priority = ConfigurationManager.AppSettings["Voc_Xtal_priority"];
                    _vocXtal_communicationType = ConfigurationManager.AppSettings["Voc_Xtal_communicationType"];
                    _vocXtal_addresseeRole = ConfigurationManager.AppSettings["Voc_Xtal_addresseeRole"];
                    _vocXtal_isColor = ConfigurationManager.AppSettings["Voc_Xtal_isColor"];
                    _vocXtal_channel = ConfigurationManager.AppSettings["Voc_Xtal_channel"];
                    _vocXtal_from = ConfigurationManager.AppSettings["Voc_Xtal_from"];
                    _vocXtal_replyTo = ConfigurationManager.AppSettings["Voc_Xtal_replyTo"];



                    //if (string.IsNullOrEmpty(_vocRootFolder) ||
                    //    string.IsNullOrEmpty(_vocFeedbackFileFolder) ||
                    //    string.IsNullOrEmpty(_vocWorkingFileFolder) ||
                    //    string.IsNullOrEmpty(_vocOutputFileFolder) ||
                    //    string.IsNullOrEmpty(_vocOutputXtalFileFolder) ||
                    //    string.IsNullOrEmpty(_vocResponseFilenamePattern) ||
                    //    string.IsNullOrEmpty(_vocBrokerFilenamePattern) ||
                    //    string.IsNullOrEmpty(_vocTemplateFileFolder))
                    //{
                    //    throw new ConfigurationErrorsException("Missing folder paths in app.config");
                    //}

                }
                catch (ConfigurationErrorsException ex)
                {
                    _log.Error($"Error reading Xtal configuration from app.config: {ex.Message}");
                    throw; // Or re-throw the exception for handling by the caller
                }
                catch (DirectoryNotFoundException ex)
                {
                    _log.Error($"Folder not found: {ex.Message}");
                    throw; // Or handle the missing folder error gracefully
                }
            }


            public  static string VocXtal_crid => _vocXtal_crid;
            public  static string VocXtal_crType => _vocXtal_crType;
            public  static string VocXtal_ownerEmail => _vocXtal_ownerEmail;
            public  static string VocXtal_ownerCostCenter => _vocXtal_ownerCostCenter;
            public  static string VocXtal_priority => _vocXtal_priority;
            public  static string VocXtal_communicationType => _vocXtal_communicationType;
            public  static string VocXtal_addresseeRole => _vocXtal_addresseeRole;
            public  static string VocXtal_isColor => _vocXtal_isColor;
            public  static string VocXtal_channel => _vocXtal_channel;
            public  static string VocXtal_from => _vocXtal_from;
            public  static string VocXtal_replyTo => _vocXtal_replyTo;
        }

        public static class JobIdGenerator
        {
            private static readonly string _jobId = GenerateUniqueJobId();

            public static string JobId => _jobId;

            private static string GenerateUniqueJobId()
            {
                string guid = Guid.NewGuid().ToString("N");
                string timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmmss");
                return string.Join("|", guid, timestamp);
            }
        }
    }
}
