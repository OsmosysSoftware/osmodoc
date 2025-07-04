using OsmoDoc.Pdf.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace OsmoDoc.Pdf;

public static class PdfDocumentGenerator
{
    /// <summary>
    /// Generates a PDF document from an HTML or EJS template.
    /// </summary>
    /// <param name="templatePath">The path to the HTML or EJS template file.</param>
    /// <param name="metaDataList">A list of content metadata to replace placeholders in the template.</param>
    /// <param name="outputFilePath">The desired output path for the generated PDF file.</param>
    /// <param name="isEjsTemplate">A boolean indicating whether the template is an EJS file.</param>
    /// <param name="serializedEjsDataJson">JSON string containing data for EJS template rendering. Required if isEjsTemplate is true.</param>
    public async static Task GeneratePdf(string templatePath, List<ContentMetaData> metaDataList, string outputFilePath, bool isEjsTemplate, string? serializedEjsDataJson)
    {
        if (metaDataList is null)
        {
            throw new ArgumentNullException(nameof(metaDataList));
        }

        if (string.IsNullOrWhiteSpace(templatePath))
        {
            throw new ArgumentNullException(nameof(templatePath));
        }

        if (string.IsNullOrWhiteSpace(outputFilePath))
        {
            throw new ArgumentNullException(nameof(outputFilePath));
        }

        if (string.IsNullOrWhiteSpace(OsmoDocPdfConfig.WkhtmltopdfPath) && !RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            throw new Exception("WkhtmltopdfPath is not set in OsmoDocPdfConfig.");
        }

        if (!File.Exists(templatePath))
        {
            throw new Exception("The file path you provided is not valid.");
        }

        if (isEjsTemplate)
        {
            // Validate if template in file path is an ejs file
            if (Path.GetExtension(templatePath).ToLower() != ".ejs")
            {
                throw new Exception("Input template should be a valid EJS file");
            }

            // Convert ejs file to an equivalent html
            templatePath = await ConvertEjsToHTML(templatePath, outputFilePath, serializedEjsDataJson);
        }

        // Modify html template with content data and generate pdf
        string modifiedHtmlFilePath = ReplaceFileElementsWithMetaData(templatePath, metaDataList, outputFilePath);
        await ConvertHtmlToPdf(OsmoDocPdfConfig.WkhtmltopdfPath, modifiedHtmlFilePath, outputFilePath);

        if (isEjsTemplate)
        {
            // If input template was an ejs file, then the template path contains path to html converted from ejs
            if (File.Exists(templatePath) && Path.GetExtension(templatePath).ToLower() == ".html")
            {
                // If template path contains path to converted html template then delete it
                File.Delete(templatePath);
            }
        }
    }

    private static string ReplaceFileElementsWithMetaData(string templatePath, List<ContentMetaData> metaDataList, string outputFilePath)
    {
        string htmlContent = File.ReadAllText(templatePath);

        foreach (ContentMetaData metaData in metaDataList)
        {
            htmlContent = htmlContent.Replace($"{{{{{metaData.Placeholder}}}}}", metaData.Content);
        }

        string? directoryPath = Path.GetDirectoryName(outputFilePath);
        if (directoryPath == null)
        {
            throw new Exception($"No directory found for the path: {outputFilePath}");
        }
        string uniqueId = Guid.NewGuid().ToString("N");
        string tempHtmlFilePath = Path.Combine(directoryPath, $"Modified_{uniqueId}");
        string tempHtmlFile = Path.Combine(tempHtmlFilePath, "modifiedHtml.html");

        if (!Directory.Exists(tempHtmlFilePath))
        {
            Directory.CreateDirectory(tempHtmlFilePath);
        }

        File.WriteAllText(tempHtmlFile, htmlContent);
        return tempHtmlFile;
    }

    private async static Task ConvertHtmlToPdf(string? wkhtmltopdfPath, string modifiedHtmlFilePath, string outputFilePath)
    {
        string fileName;
        string arguments;

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            fileName = "wkhtmltopdf";
            arguments = $"\"{modifiedHtmlFilePath}\" \"{outputFilePath}\"";
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            if (string.IsNullOrWhiteSpace(wkhtmltopdfPath))
            {
                throw new Exception("wkhtmltopdf path must be explicitly set on Windows.");
            }

            string fullPath = Path.GetFullPath(wkhtmltopdfPath);
            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"wkhtmltopdf binary not found at: {fullPath}");
            }

            fileName = fullPath;
            arguments = $"\"{modifiedHtmlFilePath}\" \"{outputFilePath}\"";
        }
        else
        {
            throw new PlatformNotSupportedException("Unsupported operating system.");
        }

        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (Process process = new Process())
        {
            process.StartInfo = psi;
            process.Start();
            await process.WaitForExitAsync();
            string output = await process.StandardOutput.ReadToEndAsync();
            string errors = await process.StandardError.ReadToEndAsync();

            if (process.ExitCode != 0)
            {
                throw new Exception($"Error during PDF generation: {errors} (Exit Code: {process.ExitCode})");
            }
        }

        // Delete the temporary modified HTML file
        if (File.Exists(modifiedHtmlFilePath))
        {
            File.Delete(modifiedHtmlFilePath);
        }
    }

    private async static Task<string> ConvertEjsToHTML(string ejsFilePath, string outputFilePath, string? ejsDataJson)
    {
        // Generate directory
        string? directoryPath = Path.GetDirectoryName(outputFilePath);
        if (directoryPath == null)
        {
            throw new Exception($"No directory found for the path: {outputFilePath}");
        }
        string uniqueId = Guid.NewGuid().ToString("N");
        string tempDirectoryFilePath = Path.Combine(directoryPath, $"Temp_{uniqueId}");

        if (!Directory.Exists(tempDirectoryFilePath))
        {
            Directory.CreateDirectory(tempDirectoryFilePath);
        }

        // Generate file path to converted html template
        string tempHtmlFilePath = Path.Combine(tempDirectoryFilePath, "htmlTemplate.html");

        // If the ejs data json is invalid then throw exception
        if (!string.IsNullOrWhiteSpace(ejsDataJson) && !IsValidJSON(ejsDataJson))
        {
            throw new Exception("Received invalid JSON data for EJS template");
        }

        // Write json data string to json file
        string ejsDataJsonFilePath = Path.Combine(tempDirectoryFilePath, "ejsData.json");
        string contentToWrite = ejsDataJson ?? "{}";
        File.WriteAllText(ejsDataJsonFilePath, contentToWrite);

        string commandLine = "cmd.exe";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            commandLine = "ejs";
        }
        string arguments = EjsToHtmlArgumentsBasedOnOS(ejsFilePath, ejsDataJsonFilePath, tempHtmlFilePath);

        ProcessStartInfo psi = new ProcessStartInfo
        {
            FileName = commandLine,
            Arguments = arguments,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using (Process process = new Process())
        {
            process.StartInfo = psi;
            process.Start();
            await process.WaitForExitAsync();
            string output = await process.StandardOutput.ReadToEndAsync();
            string errors = await process.StandardError.ReadToEndAsync();

            if (process.ExitCode != 0)
            {
                throw new Exception($"Error during EJS to HTML conversion: {errors} (Exit Code: {process.ExitCode})");
            }
        }

        // Delete json data file
        if (File.Exists(ejsDataJsonFilePath))
        {
            File.Delete(ejsDataJsonFilePath);
        }

        return tempHtmlFilePath;
    }

    private static bool IsValidJSON(string json)
    {
        try
        {
            JToken.Parse(json);
            return true;
        }
        catch (JsonReaderException)
        {
            return false;
        }
    }

    private static string EjsToHtmlArgumentsBasedOnOS(string ejsFilePath, string ejsDataJsonFilePath, string tempHtmlFilePath)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return $"/C ejs \"{ejsFilePath}\" -f \"{ejsDataJsonFilePath}\" -o \"{tempHtmlFilePath}\"";
        }
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            return $"\"{ejsFilePath}\" -f \"{ejsDataJsonFilePath}\" -o \"{tempHtmlFilePath}\"";
        }
        else
        {
            throw new Exception("Unknown operating system");
        }
    }
}
