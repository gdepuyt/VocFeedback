using System;
using System.Collections.Generic;
using static VocPoc.VocUtilsExcel;
namespace VocPoc
{

    public class VocMain
    {
        static void Main(string[] args)
        {
            AZ_BNL_StepTerminationCode stepTerminationCode = AZ_BNL_StepTerminationCode.Success;
            try
            {
                //ocSteps steps = new VocSteps();

                // Step 1: Initialize (add logging and configuration checks)
                VocSteps.Initialize();

                // Step 2: Collect and Copy Required Files
                VocSteps.CollectAndCopyRequiredFilesToWorkingFolder();

                // Step 3: Load and Filter Feedback Files
                (List<AZ_BNL_Responses_Claims> claimsList, List<AZ_BNL_Responses_Issues> issuesList,
                 List<AZ_BNL_Responses_Sales> salesList, List<AZ_BNL_Brokers> brokersList) =
                  VocSteps.LoadAndFilterFeedbackFilesInMemory(VocConfigurationManagement.FolderConfig.VocWorkingFileFolder);

                // Step 4: Merge and Write Master Excel List (handle empty data case)
                try
                {
                    VocSteps.MergeAndWriteMasterExcelList(VocConfigurationManagement.FolderConfig.VocWorkingFileFolder, claimsList, issuesList, salesList, brokersList);
                }
                catch (VocExceptionsManagement ex) when (ex.StepTerminationCode == AZ_BNL_StepTerminationCode.StoppedNonCritical)
                {
                    // Log a message about skipping file generation due to empty data
                    Console.WriteLine("No data found in any list. Skipping file generation.");
                }

                // Step 5: Get Properties from XLS Templates
                (List<PropertyInfo> configClaims, List<PropertyInfo> configIssues, List<PropertyInfo> configSales) =
                  VocSteps.GetPropertiesFromXlsTemplates();

                // Step 6: Process and Return Lists
                (List<AZ_BNL_Record> reqfields_claimsList, List<AZ_BNL_Record> reqfields_issuesList, List<AZ_BNL_Record> reqfields_salesList) =
                  VocSteps.ProcessAndReturnLists(claimsList, configClaims, issuesList, configIssues, salesList, configSales);

                // Step 7: Write Grouped Data for Types
                VocSteps.WriteGroupedDataForTypes(reqfields_claimsList, reqfields_issuesList, reqfields_salesList, VocConfigurationManagement.FolderConfig.VocTemplateFileFolder);

                // Step 8: Group All Excel Created by BCAB
                VocSteps.GroupAllExcelCreatedByBCAB();

                // Step 9: Generate Communication Request
                string communicationRequestXml = VocSteps.GenerateCommunicationRequest(VocConfigurationManagement.FolderConfig.VocOutputFileFolder);

                Console.WriteLine("Process completed successfully.");
            }
            catch (VocExceptionsManagement ex)
            {
                // Handle other exceptions with logging and potential termination

                VocExceptionsManagement.LogBatchProcessException(ex);
                stepTerminationCode = ex.StepTerminationCode;
            }
            finally
            {
                ExitBatch(stepTerminationCode);
            }
        }

        public static void ExitBatch(AZ_BNL_StepTerminationCode stepTerminationCode)
        {
            if (stepTerminationCode != AZ_BNL_StepTerminationCode.StoppedError)
            {
                Environment.Exit(0);
            }
            else { Environment.Exit(1); }
        }
    }
}

