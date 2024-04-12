using MyTestEpPLus;
using System;
namespace VocPoc
{
    /// <summary>
    /// This class provides a custom exception type for managing exceptions related to steps in a process.
    /// </summary>
    public class VocExceptionsManagement : Exception
    {
        /// <summary>
        /// Gets the step number where the exception occurred.
        /// </summary>
        public int StepNumber { get; private set; }
        /// <summary>
        /// Gets the message describing the exception related to the step.
        /// </summary>
        public string StepMessage { get; private set; }
        /// <summary>
        /// Gets the termination code associated with the exception.
        /// </summary>
        public AZ_BNL_StepTerminationCode StepTerminationCode { get; private set; }
        /// <summary>
        /// Initializes a new instance of the VocExceptionsManagement class.
        /// </summary>
        /// <param name="stepNumber">The step number where the exception occurred.</param>
        /// <param name="message">The message describing the exception.</param>
        /// <param name="stepTerminationCode">The termination code associated with the exception.</param>
        /// <param name="innerException">An optional inner exception that provides more details about the cause of the exception.</param>
        public VocExceptionsManagement(int stepNumber, string message, AZ_BNL_StepTerminationCode stepTerminationCode, Exception innerException)
            : base(message, innerException)
        {
            StepNumber = stepNumber;
            StepMessage = message;
            StepTerminationCode = stepTerminationCode;
        }
        public static void LogBatchProcessException(VocExceptionsManagement ex)
        {
            switch (ex.StepTerminationCode)
            {
                case AZ_BNL_StepTerminationCode.StoppedNonCritical:
                    VocLogger.LogStep(ex.StepNumber, false, "Batch process ended with noncritical error", VocLogger.LogLevel.Info);
                    if (ex.InnerException != null)
                    {
                        Exception currentException = ex.InnerException;
                        do
                        {
                            VocLogger.LogStep(ex.StepNumber, false, $"Exception Details :  ({currentException.GetType().Name}): {currentException.Message}", VocLogger.LogLevel.Info);
                            // Log stack trace if needed: VocLogger.LogStep(ex.StepNumber, false, currentException.StackTrace, VocLogger.LogLevel.Error);
                            currentException = currentException.InnerException;
                        } while (currentException != null);
                    }
                    else
                    {
                        // Handle exceptions without inner exceptions (optional)
                        VocLogger.LogStep(ex.StepNumber, false, ex.Message, VocLogger.LogLevel.Info);
                    }
                    break;
                case AZ_BNL_StepTerminationCode.StoppedError:
                    VocLogger.LogStep(ex.StepNumber, false, "Batch process ended with critical error", VocLogger.LogLevel.Error);
                    if (ex.InnerException != null)
                    {
                        Exception currentException = ex.InnerException;
                        do
                        {
                            VocLogger.LogStep(ex.StepNumber, false, $"Exception Details :  ({currentException.GetType().Name}): {currentException.Message}", VocLogger.LogLevel.Error);
                            // Log stack trace if needed: VocLogger.LogStep(ex.StepNumber, false, currentException.StackTrace, VocLogger.LogLevel.Error);
                            currentException = currentException.InnerException;
                        } while (currentException != null);
                    }
                    else
                    {
                        // Handle exceptions without inner exceptions (optional)
                        VocLogger.LogStep(ex.StepNumber, false, ex.Message, VocLogger.LogLevel.Error);
                    }
                    break;
                default:
                    VocLogger.LogStep(ex.StepNumber, false, ex.Message, VocLogger.LogLevel.Warn); ;
                    break;
            }
        }
    }
}
