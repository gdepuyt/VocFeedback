# VocPoc## VocPoc - A C# Application for Processing Feedback Data from InMoment

This is the repository for VocPoc, a C# application that processes feedback data related to claims, issues, and sales feedback data. 
The processed data is then used to generate reports in Excel format and potentially initiate communication requests.

### Prerequisites

* **.NET Framework 4.8 (or later):** [https://dotnet.microsoft.com/download/dotnet-framework](https://dotnet.microsoft.com/download/dotnet-framework) This is required to run the application.
* **The following NuGet packages (for managing dependencies):**

    * **CsvHelper (v31.0.2 or later):** [https://www.nuget.org/packages/CsvHelper](https://www.nuget.org/packages/CsvHelper) Used for working with CSV files.
    * **EPPlus (v7.0.10 or later):** [https://www.nuget.org/packages/EPPlus/](https://www.nuget.org/packages/EPPlus/) Used for working with Excel spreadsheets.
    * **log4net (v2.0.16 or later):** [https://www.nuget.org/packages/log4net/](https://www.nuget.org/packages/log4net/) Used for application logging.

### Features

* Collects feedback data in CSV format available in Zeus/A/....
* Filters and processes data for claims, issues, and sales.
* Generates reports in Excel format using templates.
* Groups reports by a specified property (e.g., agency ID).
* Optionally initiates a communication request based on processed data (implementation details might depend on external libraries).

### Getting Started


1. **Clone the repository:** Use `git clone https://<github_url>/VocPoc.git` to clone this repository locally.
2. **Restore NuGet packages:** Open the solution file (VocPoc.sln) in Visual Studio and ensure NuGet packages are restored for the project.
3. **Configuration:** 
    * Update the `app.ini` files with your specific folder locations for source data, templates, output, etc.
4. **Run the application:** Build and run the project in Visual Studio.


### Usage

The application processes feedback data automatically upon execution. It follows these general steps:

1. Initializes logging and configuration settings.
2. Collects feedback files from a source folder and copies them to a working folder defined with a unique identifier for the run.
3. Loads and filters data from the CSV files, separating claims, issues, and sales data.
4. Merges and writes different master Excel lists (if data exists) for claims, issues, and sales.
5. Extracts properties (information that will be extracted to me be avaiable in the excel sheet created and collected later on) from template Excel files.
6. Processes the data lists using the extracted properties.
7. Writes grouped data for claims, issues, and sales to separate Excel reports using designated folders. Gro
8. Groups all generated Excel files per borkers (BCAB).
9. Generates a communication request.


### Project Structure

* **main.cs:** This file contains the main entry point for the application.
* **vocsteps.cs:** This file defines classes and methods for processing feedback data and generating reports.
* **Other files:** The project might contain additional helper classes, configuration files, and utility functions depending on the implementation.


### Folder Structure

The application uses the following folder structure:

**Root Directory (VocPoc/)**

* **Archives** (Folder): Stores **original feedback files received from InMoment**. These files serve as the source data for processing within VocPoc. Subfolders are typically named by date (YYYYMMDD).  
    * Example: `IM2AZBNL/20240311` (Folder containing original feedback files from InMoment received on March 11th, 2024)
* **Feedbacks** (Folder): Contains folders for feedback data processing.
    * **Dependencies** (Folder): Stores shared resources used by the application.
        * **BrokersList** (Folder): Contains data related to brokers  
        * **Templates** (Folder): Stores Excel templates used for report generation.
        * **Translations** (Folder): Contain translation files for French/Dutch support 
    * **Output** (Folder): Contains the generated reports in Excel format. Subfolders are typically named based on report type (Claims, Issues, Sales). Also include an `XtalCR` folder to store Communication request used by Xtal to send email to brokers
        * Example: `Output/Claims/202404120546-4b6a6d19a37649a397ba303cc9c4cd07` (Folder for claims report generated on 2024-04-12 with unique identifier)
    * **WorkingFolder** (Folder): Used as a temporary workspace during processing. Subfolders are typically named with a unique identifier.
        * Example: `WorkingFolder/202404120546-4b6a6d19a37649a397ba303cc9c4cd07` (Temporary workspace for processing data)
* **Logs** (Folder): Stores application log files.

```