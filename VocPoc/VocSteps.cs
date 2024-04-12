using MyTestEpPLus;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using static VocPoc.VocUtilsCommon;
using static VocPoc.VocUtilsExcel;

namespace VocPoc
{
    internal static class VocSteps
    {
        public static void Initialize()
        {
            try
            {
                // Configure logging
                VocLogger.LogStep(1, true, "Invoke Initialize method", VocLogger.LogLevel.Info);
                VocLogger.LogStep(1, false, $"JobID for this run = {VocConfigurationManagement.JobIdGenerator.JobId}", VocLogger.LogLevel.Info);
                log4net.Config.XmlConfigurator.Configure();
                // Check and log configuration settings
                VocConfigurationManagement.CheckAndLogAPPSettings();
                // Validate and retrieve folder paths
                string rootFolder = VocConfigurationManagement.FolderConfig.VocRootFolder;
                // ... access other folder paths
                //Check if translations & brokers files are availalbe and well formated
                CsvValidator.ValidateCsv(Path.Combine(VocConfigurationManagement.FolderConfig.VocTranslationsFileFolder, VocConfigurationManagement.FilePatternConfig.VocTranslationsFilePattern), typeof(AZ_BNL_FieldMapping));
                CsvValidator.ValidateCsv(Path.Combine(VocConfigurationManagement.FolderConfig.VocBrokerFileFolder, VocConfigurationManagement.FilePatternConfig.VocBrokerFilenamePattern), typeof(AZ_BNL_Brokers));
                VocFileManagement.FileExists(Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_IssuesFilenamePattern));
                VocFileManagement.FileExists(Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_ClaimsFilenamePattern));
                VocFileManagement.FileExists(Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_SalesFilenamePattern));
            }
            catch (Exception ex)
            {
                throw new VocExceptionsManagement(1, "Error during the initialisation - batch is stopped", AZ_BNL_StepTerminationCode.StoppedError, ex);
            }
        }
        public static void CollectAndCopyRequiredFilesToWorkingFolder()
        {
            try
            {
                VocLogger.LogStep(2, true, "Invoke CollectAndCopyRequiredFilesToWorking method", VocLogger.LogLevel.Info);
                // Call the CopyAndRenameFiles method for feedback files

                VocFileManagement.CopyAndRenameFiles(
                  VocConfigurationManagement.FolderConfig.VocFeedbackFileFolder,
                  VocConfigurationManagement.FolderConfig.VocWorkingFileFolder,
                  VocConfigurationManagement.JobIdGenerator.JobId,
                  VocConfigurationManagement.FilePatternConfig.VocResponseFilenamePattern
                );
                // Call the CopyAndRenameFiles method for broker file ithout renaming the source file after the copy.
                VocFileManagement.CopyAndRenameFiles(
                  VocConfigurationManagement.FolderConfig.VocBrokerFileFolder,
                  VocConfigurationManagement.FolderConfig.VocWorkingFileFolder,
                  VocConfigurationManagement.JobIdGenerator.JobId,
                  VocConfigurationManagement.FilePatternConfig.VocBrokerFilenamePattern,
                  false
                );
                // Call the CopyAndRenameFiles method for translations file without renaming the source file after the copy.
                VocFileManagement.CopyAndRenameFiles(
                  VocConfigurationManagement.FolderConfig.VocTranslationsFileFolder,
                  VocConfigurationManagement.FolderConfig.VocWorkingFileFolder,
                  VocConfigurationManagement.JobIdGenerator.JobId,
                  VocConfigurationManagement.FilePatternConfig.VocTranslationsFilePattern,
                  false
                );
            }
            catch (Exception ex)
            {
                throw new VocExceptionsManagement(2,"No files available corresponding to the search criteria", AZ_BNL_StepTerminationCode.StoppedNonCritical, ex);
            }
        }
        public static (List<AZ_BNL_Responses_Claims>, List<AZ_BNL_Responses_Issues>, List<AZ_BNL_Responses_Sales>, List<AZ_BNL_Brokers>) LoadAndFilterFeedbackFilesInMemory(string workingFolder)
        {
            //_log.Info("STEP03. Load all responses files and map them towards differents lists.");
            VocLogger.LogStep(3, true, "LoadAndFilterFeedbackFilesInMemor", VocLogger.LogLevel.Info);
            // Read all CSV files from the working folder
            List<object> allData = ReadAllCsvFilesFromFolder(workingFolder);
            // Filter data based on type and criteria
            var claimsList = allData.OfType<AZ_BNL_Responses_Claims>()
                                     .Where(obj => !obj.UID.Contains("_NL_"))
                                     .ToList();
            var issueList = allData.OfType<AZ_BNL_Responses_Issues>()
                                    .Where(obj => !obj.UID.Contains("_NL_"))
                                     .ToList();
            var salesList = allData.OfType<AZ_BNL_Responses_Sales>()
                                    .Where(obj => !obj.UID.Contains("_NL_"))
                                     .ToList();
            var brokersList = allData.OfType<AZ_BNL_Brokers>()
                                     .ToList();
            // Return all filtered lists as a tuple
            return (claimsList, issueList, salesList, brokersList);
        }
        public static void MergeAndWriteMasterExcelList(string workingFolder, List<AZ_BNL_Responses_Claims> claimsList, List<AZ_BNL_Responses_Issues> issueList, List<AZ_BNL_Responses_Sales> salesList, List<AZ_BNL_Brokers> brokersList)
        {
            VocLogger.LogStep(4, true, "MergeAndWriteMasterExcelList is starting", VocLogger.LogLevel.Info);
            if (!claimsList.Any() && !issueList.Any() && !salesList.Any())
            {
                //_log.Info("No data found in any list. Skipping file generation.");
                //throw new DataException("No data found in any list. Exiting program."); // Throw an exception
                throw new VocExceptionsManagement(4, "No data found in any list. Skipping file generation", AZ_BNL_StepTerminationCode.StoppedNonCritical, new Exception("Claims, Issue or Sales list are empty"));
            }
            else
            {
                try
                {
                    // Write claims list (check if list is not empty)
                    if (claimsList.Any())
                    {
                        VocUtilsExcel.WriteObjectsToExcel<AZ_BNL_Responses_Claims>(
                          Path.Combine(workingFolder, VocFileManagement.AddUniqueStampToFileName("MasterClaimsList.xlsx")), claimsList);
                    }
                    // Write issue list (check if list is not empty)
                    if (issueList.Any())
                    {
                        VocUtilsExcel.WriteObjectsToExcel<AZ_BNL_Responses_Issues>(
                          Path.Combine(workingFolder, VocFileManagement.AddUniqueStampToFileName("MasterIssueList.xlsx")), issueList);
                    }
                    // Write sales list (check if list is not empty)
                    if (salesList.Any())
                    {
                        VocUtilsExcel.WriteObjectsToExcel<AZ_BNL_Responses_Sales>(
                          Path.Combine(workingFolder, VocFileManagement.AddUniqueStampToFileName("MasterSalesList.xlsx")), salesList);
                    }
                    // Write brokers list (check if list is not empty)
                    if (brokersList.Any())
                    {
                        VocUtilsExcel.WriteObjectsToExcel<AZ_BNL_Brokers>(
                          Path.Combine(workingFolder, VocFileManagement.AddUniqueStampToFileName("MasterBrokersList.xlsx")), brokersList);
                    }
                }
                //}
                catch (Exception ex)
                {
                    throw new VocExceptionsManagement(4, "Error during the creation of the masterlists for claims, issue,sales in  step 4", AZ_BNL_StepTerminationCode.StoppedError, ex);
                }
            }
        }
        public static (List<PropertyInfo> ConfigClaims, List<PropertyInfo> ConfigIssue, List<PropertyInfo> ConfigSales) GetPropertiesFromXlsTemplates()
        {
            VocLogger.LogStep(5, true, "GetPropertiesFromXlsTemplates", VocLogger.LogLevel.Info);
            try
            {
                // Fetch properties for each template
                List<PropertyInfo> ConfigClaims = VocUtilsExcel.GetPropertiesInRange(
                  Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_ClaimsFilenamePattern),
                  VocConfigurationManagement.RangeConfig.VocXlsTemplate_ClaimsRange);
                List<PropertyInfo> ConfigIssue = VocUtilsExcel.GetPropertiesInRange(
                  Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_IssuesFilenamePattern),
                  VocConfigurationManagement.RangeConfig.VocXlsTemplate_IssuesRange);
                List<PropertyInfo> ConfigSales = VocUtilsExcel.GetPropertiesInRange(
                  Path.Combine(VocConfigurationManagement.FolderConfig.VocTemplateFileFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_SalesFilenamePattern),
                  VocConfigurationManagement.RangeConfig.VocXlsTemplate_SalesRange);
                return (ConfigClaims, ConfigIssue, ConfigSales);
            }
            catch (Exception ex)
            {
                throw new VocExceptionsManagement(5, "Error identifying properties in XLS templates", AZ_BNL_StepTerminationCode.StoppedError, ex);
            }
        }
        public static (List<AZ_BNL_Record>, List<AZ_BNL_Record>, List<AZ_BNL_Record>) ProcessAndReturnLists(
            List<AZ_BNL_Responses_Claims> claimsList, List<PropertyInfo> configClaims,
            List<AZ_BNL_Responses_Issues> issuesList, List<PropertyInfo> configIssues,
            List<AZ_BNL_Responses_Sales> salesList, List<PropertyInfo> configSales)
        {
            VocLogger.LogStep(6, true, "ProcessAndReturnLists", VocLogger.LogLevel.Info);
            try
            {
                List<AZ_BNL_Record> reqfields_claimsList = GetSelectedAttributes(claimsList, configClaims);
                List<AZ_BNL_Record> reqfields_issuesList = GetSelectedAttributes(issuesList, configIssues);
                List<AZ_BNL_Record> reqfields_salesList = GetSelectedAttributes(salesList, configSales);
                // Use var for type inference (compiler will infer the type as a tuple)
                return (reqfields_claimsList, reqfields_issuesList, reqfields_salesList);
            }
            catch (Exception ex)
            {
                //_log.Error("Error during list processing: {0}");
                //_log.Error($"{ex.Message}");
                throw new VocExceptionsManagement(6, "Error during list processing", AZ_BNL_StepTerminationCode.StoppedError, ex);
            }
        }
        public static void WriteGroupedDataForTypes(
            List<AZ_BNL_Record> claimsList,
            List<AZ_BNL_Record> issuesList,
            List<AZ_BNL_Record> salesList,
            string templateFolder
        )
        {
            VocLogger.LogStep(7, true, "WriteGroupedDataForTypes", VocLogger.LogLevel.Info);
            try
            {
                // Claims
                string claimsTemplatePath = Path.Combine(templateFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_ClaimsFilenamePattern);
                VocUtilsExcel.WriteGroupedDataToExcel(
                    VocConfigurationManagement.FolderConfig.VocOutputClaimsFileFolder,
                    claimsList,
                    obj => obj.CustSales_AgencyID,
                    claimsTemplatePath,
                    "Claims"
                );
                // Issues
                string issuesTemplatePath = Path.Combine(templateFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_IssuesFilenamePattern);
                VocUtilsExcel.WriteGroupedDataToExcel(
                    VocConfigurationManagement.FolderConfig.VocOutputIssuesFileFolder,
                    issuesList,
                    obj => obj.CustSales_AgencyID,
                    issuesTemplatePath,
                    "Issues"
                );
                // Sales
                string salesTemplatePath = Path.Combine(templateFolder, VocConfigurationManagement.FilePatternConfig.VocXlsTemplate_SalesFilenamePattern);
                VocUtilsExcel.WriteGroupedDataToExcel(
                    VocConfigurationManagement.FolderConfig.VocOutputSalesFileFolder,
                    salesList,
                    obj => obj.CustSales_AgencyID,
                    salesTemplatePath,
                    "Sales"
                );
            }
            catch (Exception ex)
            {
                //_log.Error("Error writing grouped data to Excel: {0}");
                //_log.Error($"{ex.Message}");
                throw new VocExceptionsManagement(7, "Identifying properties defined in the XLS templates for reporting...", AZ_BNL_StepTerminationCode.StoppedError, ex);
                // Consider additional actions, such as logging more details or retrying
            }
        }
        public static void GroupAllExcelCreatedByBCAB()
        {
            VocLogger.LogStep(8, true, "GroupAllExcelCreatedByBCAB", VocLogger.LogLevel.Info);
            string[] folderPaths = new string[]
            {
                VocConfigurationManagement.FolderConfig.VocOutputClaimsFileFolder,
                VocConfigurationManagement.FolderConfig.VocOutputIssuesFileFolder,
                VocConfigurationManagement.FolderConfig.VocOutputSalesFileFolder
            };
            VocUtilsExcel.FileGrouper.GroupFiles(folderPaths);
        }
        public static string GenerateCommunicationRequest(string outputFolder)
        {
            VocLogger.LogStep(9, true, "GenerateCommunicationRequest", VocLogger.LogLevel.Info);
            return XtalUtils.CRRequestGenerator.CRRequestInit(outputFolder);
        }

    }
}