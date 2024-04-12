using MyTestEpPLus;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Reflection;

namespace VocPoc
{
    /// <summary>
    /// This class provides methods for managing configuration settings used by the Voc Feedback application.
    /// </summary>
    internal class VocConfigurationManagement
    {
        /// <summary>
        /// Reads application settings from app.config, logs them, and validates required configuration settings.
        /// 
        /// Throws ConfigurationErrorsException if a required  configuration property is missing or empty.
        /// </summary>
        public static class ConfigurationHelper
        {
            private static readonly Dictionary<string, Type> ClassTypeMapping = new Dictionary<string, Type>()
              {
                { "Xtal", typeof(XtalConfig) },
                { "Range", typeof(RangeConfig) },
                { "FilePattern", typeof(FilePatternConfig) }
              };


            /// <summary>
            /// Validates if required string properties are set for a specific configuration class (`Xtal`, `Range`, `FilePattern`).
            /// 
            /// Throws ConfigurationErrorsException if a required string property is missing or empty.
            /// </summary>
            /// <param name="className">The name of the configuration class.</param>
            /// <param name="configurationSource">The configuration source (optional, default: "app.config").</param>         
            public static void ValidateRequiredStringProperties(string className, string configurationSource = "app.config")
            {
                Type classType = GetClassType(className);

                var fields = classType.GetFields(BindingFlags.Static | BindingFlags.NonPublic);

                foreach (var field in fields)
                {
                    //if (field.Name.StartsWith("_vocXtal") && field.FieldType == typeof(string))
                    //{
                        string value = (string)field.GetValue(null);
                        if (string.IsNullOrEmpty(value))
                        {
                            throw new ConfigurationErrorsException($"Missing or empty value for configuration setting: {field.Name} ({configurationSource})");
                        }
                    //}
                }
            }

            private static Type GetClassType(string className)
            {
                if (ClassTypeMapping.TryGetValue(className, out Type classType))
                {
                    return classType;
                }
                else
                {
                    throw new ArgumentException($"Class name '{className}' not found in mapping");
                }
            }
        }

        /// <summary>
        /// Reads application settings from app.config, logs them, and validates required configuration settings.
        /// 
        /// Throws ConfigurationErrorsException if a required configuration setting is missing or empty.
        /// </summary>
        public static void CheckAndLogAPPSettings()
        {
            try
            {
                // Read all application settings from app.config
                var appSettings = ConfigurationManager.AppSettings;
                
                // Log all settings with key-value pairs
                foreach (var key in appSettings.AllKeys)
                {
                    string value = appSettings[key];

                    VocLogger.LogStep(1, false, $"App setting: {key} = {value}", VocLogger.LogLevel.Info);
                   
                }

                // Validate required configuration settings
                ConfigurationHelper.ValidateRequiredStringProperties("Xtal");
                ConfigurationHelper.ValidateRequiredStringProperties("Range");
                ConfigurationHelper.ValidateRequiredStringProperties("FilePattern");



            }
            catch (ConfigurationErrorsException ex)
            {
                // Re-throw for handling by caller
                throw ex;

            }
            catch (Exception ex)
            {
                //Re-throw for handling by caller
                throw ex;

            }
        }

        public static class FolderConfig
        {
            private static string _vocRootFolder;
            private static string _vocFeedbackFileFolder;
            private static string _vocWorkingFileFolder;
            private static string _vocOutputFileFolder;
            private static string _vocOutputXtalFileFolder;
            private static string _vocOutputClaimsFileFolder;
            private static string _vocOutputIssuesFileFolder;
            private static string _vocOutputSalesFileFolder;
            private static string _vocBrokerFileFolder;
            private static string _vocTemplateFileFolder;
            private static string _vocTranslationsFileFolder;

            static FolderConfig()
            {
                try
                {
                    string[] folderPaths = {
                        "VocRootFolder",
                        "VocFeedbackFileFolder",
                        "VocWorkingFileFolder",
                        "VocOutputFileFolder",
                        "VocOutputClaimsFileFolder",
                        "VocOutputIssuesFileFolder",
                        "VocOutputSalesFileFolder",
                        "VocOutputXtalFileFolder",
                        "VocTemplateFileFolder",
                        "VocBrokerFileFolder",
                        "VocTranslationsFileFolder"

                    };


                    // Read folder paths from app.config
                    Dictionary<string, string> folderMap = new Dictionary<string, string>();
                    foreach (string key in folderPaths)
                    {
                        folderMap[key] = ConfigurationManager.AppSettings[key];
                    }

                    
                    foreach (KeyValuePair<string, string> entry in folderMap)
                    {
                        if (string.IsNullOrEmpty(entry.Value))
                        {
                            throw new ConfigurationErrorsException($"Missing folder path '{entry.Key}' in app.config");
                        }

                        if (!Directory.Exists(entry.Value))
                        {
                                                        throw new ConfigurationErrorsException($"Missing folder path '{entry.Value}' on the server");

                        }
                    }

                   
                    // Assign folder paths to properties (assuming successful validation)
                    _vocRootFolder = folderMap["VocRootFolder"];
                    _vocFeedbackFileFolder = folderMap["VocFeedbackFileFolder"];
                    _vocWorkingFileFolder = Path.Combine(folderMap["VocWorkingFileFolder"], VocConfigurationManagement.JobIdGenerator.JobId);
                    _vocOutputFileFolder = Path.Combine(folderMap["VocOutputFileFolder"], VocConfigurationManagement.JobIdGenerator.JobId);
                    _vocOutputXtalFileFolder = folderMap["VocOutputXtalFileFolder"];
                    _vocOutputClaimsFileFolder = Path.Combine(folderMap["VocOutputClaimsFileFolder"], VocConfigurationManagement.JobIdGenerator.JobId);
                    _vocOutputIssuesFileFolder = Path.Combine(folderMap["VocOutputIssuesFileFolder"], VocConfigurationManagement.JobIdGenerator.JobId);
                    _vocOutputSalesFileFolder = Path.Combine(folderMap["VocOutputSalesFileFolder"], VocConfigurationManagement.JobIdGenerator.JobId);
                    _vocTemplateFileFolder = folderMap["VocTemplateFileFolder"];
                    _vocBrokerFileFolder = folderMap["VocBrokerFileFolder"];
                    _vocTranslationsFileFolder = folderMap["VocTranslationsFileFolder"];

                    // ... assign other folder properties
                }
                catch (ConfigurationErrorsException)
                {
                    //_log.Error($"Error reading folder paths from app.config: {ex.Message}");
                    throw;
                    //Environment.Exit(1); // Exit with an error code
                }
            }




            public static string VocRootFolder => _vocRootFolder;
            public static string VocFeedbackFileFolder => _vocFeedbackFileFolder;
            public static string VocWorkingFileFolder => _vocWorkingFileFolder;
            public static string VocOutputFileFolder => _vocOutputFileFolder;
            public static string VocOutputXtalFileFolder => _vocOutputXtalFileFolder;
            public static string VocOutputClaimsFileFolder => _vocOutputClaimsFileFolder;
            public static string VocOutputIssuesFileFolder => _vocOutputIssuesFileFolder;
            public static string VocOutputSalesFileFolder => _vocOutputSalesFileFolder;
            public static string VocTemplateFileFolder => _vocTemplateFileFolder;
            public static string VocBrokerFileFolder => _vocBrokerFileFolder;
            public static string VocTranslationsFileFolder => _vocTranslationsFileFolder;
        }

        public static class FilePatternConfig
        {
            private static string _vocResponseFilenamePattern;
            private static string _vocBrokerFilenamePattern;
            private static string _vocBrokerOptoutFilenamePattern;
            private static string _vocTranslationsFilenamePattern;
            private static string _voc_XlsTemplate_IssuesFilenamePattern;
            private static string _voc_XlsTemplate_ClaimsFilenamePattern;
            private static string _voc_XlsTemplate_SalesFilenamePattern;



            static FilePatternConfig()
            {

                _vocResponseFilenamePattern = ConfigurationManager.AppSettings["VocResponseFilenamePattern"];
                _vocBrokerFilenamePattern = ConfigurationManager.AppSettings["VocBrokerFilenamePattern"];
                _vocBrokerOptoutFilenamePattern = ConfigurationManager.AppSettings["VocBrokerOptoutFilenamePattern"];
                _vocTranslationsFilenamePattern = ConfigurationManager.AppSettings["VocTranslationsFilenamePattern"];
                _voc_XlsTemplate_IssuesFilenamePattern = ConfigurationManager.AppSettings["Voc_XlsTemplate_IssuesFilenamePattern"];
                _voc_XlsTemplate_ClaimsFilenamePattern = ConfigurationManager.AppSettings["Voc_XlsTemplate_ClaimsFilenamePattern"];
                _voc_XlsTemplate_SalesFilenamePattern = ConfigurationManager.AppSettings["Voc_XlsTemplate_SalesFilenamePattern"];

            }

            public static string VocResponseFilenamePattern => _vocResponseFilenamePattern;
            public static string VocBrokerFilenamePattern => _vocBrokerFilenamePattern;
            public static string VocBrokerOptOutFilenamePattern => _vocBrokerOptoutFilenamePattern;
            public static string VocTranslationsFilePattern => _vocTranslationsFilenamePattern;
            public static string VocXlsTemplate_IssuesFilenamePattern => _voc_XlsTemplate_IssuesFilenamePattern;
            public static string VocXlsTemplate_ClaimsFilenamePattern => _voc_XlsTemplate_ClaimsFilenamePattern;
            public static string VocXlsTemplate_SalesFilenamePattern => _voc_XlsTemplate_SalesFilenamePattern;


        }


        public static class RangeConfig
        {

            private static string _voc_XlsTemplate_IssuesRange;
            private static string _voc_XlsTemplate_ClaimsRange;
            private static string _voc_XlsTemplate_SalesRange;



            static RangeConfig()
            {


                _voc_XlsTemplate_IssuesRange = ConfigurationManager.AppSettings["Voc_XlsTemplate_IssuesRange"];
                _voc_XlsTemplate_ClaimsRange = ConfigurationManager.AppSettings["Voc_XlsTemplate_ClaimsRange"];
                _voc_XlsTemplate_SalesRange = ConfigurationManager.AppSettings["Voc_XlsTemplate_SalesRange"];

            }

            public static string VocXlsTemplate_IssuesRange => _voc_XlsTemplate_IssuesRange;
            public static string VocXlsTemplate_ClaimsRange => _voc_XlsTemplate_ClaimsRange;
            public static string VocXlsTemplate_SalesRange => _voc_XlsTemplate_SalesRange;


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
                                               
            }


            public static string VocXtal_crid => _vocXtal_crid;
            public static string VocXtal_crType => _vocXtal_crType;
            public static string VocXtal_ownerEmail => _vocXtal_ownerEmail;
            public static string VocXtal_ownerCostCenter => _vocXtal_ownerCostCenter;
            public static string VocXtal_priority => _vocXtal_priority;
            public static string VocXtal_communicationType => _vocXtal_communicationType;
            public static string VocXtal_addresseeRole => _vocXtal_addresseeRole;
            public static string VocXtal_isColor => _vocXtal_isColor;
            public static string VocXtal_channel => _vocXtal_channel;
            public static string VocXtal_from => _vocXtal_from;
            public static string VocXtal_replyTo => _vocXtal_replyTo;
        }

        /// <summary>
        /// Gets a unique job ID generated on class initialization.
        /// </summary>
        public static class JobIdGenerator
        {
            private static readonly string _jobId = GenerateUniqueJobId();

            public static string JobId => _jobId;

            private static string GenerateUniqueJobId()
            {
                string guid = Guid.NewGuid().ToString("N");
                string timestamp = DateTime.UtcNow.ToString("yyyyMMddHHmm");
                return string.Join("-", timestamp, guid);
            }
        }
       }
}


