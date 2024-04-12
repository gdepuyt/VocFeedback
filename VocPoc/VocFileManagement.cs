using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace VocPoc
{
    /// <summary>
    /// This static class provides utility methods for managing files.
    /// </summary>
    public static class VocFileManagement
    {
        /// <summary>
        /// Gets a list of file paths that match a specified response file name pattern within a directory and its subdirectories.
        /// </summary>
        /// <param name="directoryPath">The path to the directory to search.</param>
        /// <param name="responseFilenamePattern">The regular expression pattern for matching response filenames.</param>
        /// <returns>A list of file paths that match the pattern.</returns>
        private static List<string> GetResponseFiles(string directoryPath, string responseFilenamePattern)
        {
            List<string> responseFiles = new List<string>();
            foreach (string filePath in Directory.EnumerateFiles(directoryPath, "*.*", SearchOption.AllDirectories))
            {
                if (IsResponseFile(filePath, responseFilenamePattern))
                {
                    responseFiles.Add(filePath);
                }
            }
            return responseFiles;
        }

        /// <summary>
        /// Checks if a file path matches a specified response file name pattern using regular expressions.
        /// </summary>
        /// <param name="filePath">The path to the file to check.</param>
        /// <param name="responseFilenamePattern">The regular expression pattern for matching response filenames.</param>
        /// <returns>True if the file path matches the pattern, false otherwise.</returns>
        private static bool IsResponseFile(string filePath, string responseFilenamePattern)
        {
            return Regex.IsMatch(filePath, responseFilenamePattern);
        }

        /// <summary>
        /// Copies and renames files based on a response file name pattern and job ID.
        /// </summary>
        /// <param name="sourceDirectory">The directory containing the source files.</param>
        /// <param name="workingFolder">The directory to copy the files to.</param>
        /// <param name="jobId">The job ID to use for generating unique filenames (optional).</param>
        /// <param name="responseFilenamePattern">The regular expression pattern for matching response filenames.</param>
        /// <param name="renameSourceFiles">A flag indicating whether to rename the source files (default: true).</param>
        public static void CopyAndRenameFiles(string sourceDirectory, string workingFolder, string jobId, string responseFilenamePattern, bool renameSourceFiles = true)
        {
            List<string> responseFiles = GetResponseFiles(sourceDirectory, responseFilenamePattern);

            // Check if the list is empty before proceeding
            if (!responseFiles.Any())
            {
                throw new ArgumentException($"No feedback files found with the specified criteria. List of files for {responseFilenamePattern} in folder {sourceDirectory} is empty.");
            }

            Directory.CreateDirectory(workingFolder); // Create job folder

            foreach (string filePath in responseFiles)
            {
                string fileName = Path.GetFileName(filePath);
                string newFilePath = Path.Combine(workingFolder, fileName);

                // Copy the file with overwrite option
                File.Copy(filePath, newFilePath, true); // Overwrites if the file already exists

                if (renameSourceFiles)
                {
                    string renamedPath = Path.Combine(Path.GetDirectoryName(filePath), AddUniqueStampToFileName(fileName));
                    File.Move(filePath, renamedPath); // Rename the original file
                }
            }
        }

        /// <summary>
        /// Adds a unique timestamp or job ID (if specified) to a filename.
        /// </summary>
        /// <param name="fileName">The original filename.</param>
        /// <param name="useJobId">A flag indicating whether to use the job ID (default: true).</param>
        /// <returns>The filename with a unique identifier appended.</returns>
        public static string AddUniqueStampToFileName(string fileName, bool useJobId = true)
        {
            string uniqueIdentifier;

            if (useJobId)
            {
                // Use JobID if specified
                uniqueIdentifier = VocConfigurationManagement.JobIdGenerator.JobId;
            }
            else
            {
                // Use timestamp by default
                uniqueIdentifier = DateTime.Now.ToString("yyyyMMdd_HHmm");
            }

            return Path.GetFileNameWithoutExtension(fileName) + "-" + uniqueIdentifier
                   + Path.GetExtension(fileName);
        }

        /// <summary>
        /// Checks if a file exists at the specified path.
        /// </summary>
        /// <param name="filePath">The path to the file to check.</param>
        /// <returns>True if the file exists, false otherwise.</returns>
        /// <exception cref="ArgumentException">Throws an Argument
        public static bool FileExists(string filePath)
        {
            try
            {
                return File.Exists(filePath);
            }
            catch (ArgumentNullException ex)
            {
                throw new ArgumentException("File path cannot be null.", ex); // Re-throw with a more specific message
            }
            catch (Exception ex)
            {
                // Log the error (consider using a logging library)
                Console.WriteLine($"Error checking file existence: {ex.Message}");
                return false; // Default to false on any other exception
            }
        }

        /// <summary>
        /// This class encapsulates information about a file.
        /// </summary>
        public class FileInfoWrapper
        {
            /// <summary>
            /// Gets the full path of the file.
            /// </summary>
            public string FullPath { get; set; }

            /// <summary>
            /// Gets the filename including the extension.
            /// </summary>
            public string FileName { get; set; }

            /// <summary>
            /// Gets the filename without the extension.
            /// </summary>
            public string FilenameWithoutExtension { get; set; }
        }

        /// <summary>
        /// Gets a list of FileInfoWrapper objects containing information about the files in a directory.
        /// </summary>
        /// <param name="folderPath">The path to the directory.</param>
        /// <returns>A list of FileInfoWrapper objects.</returns>
        /// <exception cref="ArgumentException">Throws an ArgumentException if the folder path doesn't exist.</exception>
        public static List<FileInfoWrapper> GetFilesInfo(string folderPath)
        {
            // Validate folder path (optional)
            if (!Directory.Exists(folderPath))
            {
                throw new ArgumentException("Folder path doesn't exist.");
            }

            // Get all files in the folder
            var files = Directory.GetFiles(folderPath);

            // Convert files to FileInfoWrapper objects
            return files.Select(filePath => new FileInfoWrapper
            {
                FullPath = filePath,
                FileName = Path.GetFileName(filePath),
                FilenameWithoutExtension = Path.GetFileNameWithoutExtension(filePath)
            }).ToList();
        }

        /// <summary>
        /// Validates if the specified folder exists and throws a DirectoryNotFoundException if not.
        /// </summary>
        /// <param name="folderPath">The path to the folder to validate.</param>
        /// <exception cref="DirectoryNotFoundException">Throws a DirectoryNotFoundException if the folder doesn't exist.</exception>
        public static void ValidateFolderExistence(string folderPath)
        {
            if (!Directory.Exists(folderPath))
            {
                throw new DirectoryNotFoundException($"Folder '{folderPath}' does not exist.");
            }
        }

        /// <summary>
        /// Converts a comma-separated string of filenames to an array of strings.
        /// </summary>
        /// <param name="commaSeparatedFilenames">The comma-separated string containing filenames.</param>
        /// <returns>An array of strings representing the individual filenames.</returns>
        public static string[] FolderStringToArray(string commaSeparatedFilenames)
        {
            if (string.IsNullOrEmpty(commaSeparatedFilenames))
            {
                return Array.Empty<string>(); // Handle empty string or null value
            }

            return commaSeparatedFilenames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
        }

    }

}


