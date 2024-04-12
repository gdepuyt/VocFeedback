using System;
using System.Collections.Generic;
using System.Linq;
using VocPoc;

namespace VocPoC
{
    /// <summary>
    /// This class likely deals with managing translation logic for the application.
    /// </summary>
    internal class VocTranslationManagement
    {
        /// <summary>
        /// Provides methods for loading and retrieving field translations.
        /// </summary>
        internal static class FieldMappingProvider
        {
            private static readonly string _translationCsvFolder;
            private static readonly string _translationCsvFile;
            private static readonly List<AZ_BNL_FieldMapping> _cachedTranslation;

            static FieldMappingProvider()
            {
                _translationCsvFolder = VocConfigurationManagement.FolderConfig.VocWorkingFileFolder;
                _translationCsvFile = VocConfigurationManagement.FilePatternConfig.VocTranslationsFilePattern;

                _cachedTranslation = LoadMappings();
            }

            /// <summary>
            /// Loads field translations from CSV files located in a specific folder based on a filename pattern.
            /// </summary>
            /// <returns>A list of AZ_BNL_FieldMapping objects containing the loaded translations.</returns>
            private static List<AZ_BNL_FieldMapping> LoadMappings()
            {
                List<AZ_BNL_FieldMapping> mappings = new List<AZ_BNL_FieldMapping>();

                try
                {
                    List<object> allTranslation = VocUtilsCommon.ReadAllCsvFilesFromFolder(_translationCsvFolder, _translationCsvFile);
                    mappings = allTranslation.OfType<AZ_BNL_FieldMapping>().ToList();
                }
                catch (Exception ex)
                {
                    // Handle general exception
                    Console.WriteLine("Error loading field mapping for translation: {0}", ex.Message);
                    throw; // Rethrow to allow further handling if needed
                }

                return mappings;
            }

            /// <summary>
            /// Retrieves the translated value for a given field name and value, optionally considering a preferred language.
            /// </summary>
            /// <param name="fieldName">The name of the field to translate.</param>
            /// <param name="fieldValue">The current value of the field.</param>
            /// <param name="defaultLanguage">The desired language for the translation (optional). Defaults to null. Valid values are likely "F" (French) and "N" (Dutch).</param>
            /// <returns>
            /// The translated value for the field in the specified language, or the original field value if no translation is found or no language is specified.
            /// </returns>
            public static string GetTranslation(string fieldName, string fieldValue, string defaultLanguage = null)
            {
                var mapping = _cachedTranslation.FirstOrDefault(m => m.FieldName.Trim() == fieldName.Trim() && m.FieldValue.Trim() == fieldValue.Trim());

                if (mapping != null)
                {
                    switch (defaultLanguage)
                    {
                        case "F":
                            return mapping.FrenchTranslation;
                        case "N":
                            return mapping.DutchTranslation;
                        default:
                            // Return default field value if no specific language requested
                            return mapping.FieldValue;
                    }
                }
                else
                {
                    // Return field value if no mapping found
                    return fieldValue;
                }
            }
        }
    }

}