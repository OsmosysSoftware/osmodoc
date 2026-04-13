using DocumentService.Pdf.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace DocumentService.Pdf
{
    public class PdfDocumentGenerator
    {
        public static void GeneratePdf(string toolFolderAbsolutePath, string templatePath, List<ContentMetaData> metaDataList, string outputFilePath, bool isEjsTemplate, string serializedEjsDataJson)
        {
            try
            {
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
                    templatePath = ConvertEjsToHTML(templatePath, outputFilePath, serializedEjsDataJson);
                }

                // Modify html template with content data and generate pdf
                string modifiedHtmlFilePath = ReplaceFileElementsWithMetaData(templatePath, metaDataList, outputFilePath);
                ConvertHtmlToPdf(toolFolderAbsolutePath, modifiedHtmlFilePath, outputFilePath);

                if (isEjsTemplate)
                {
                    // If input template was an ejs file, then the template path contains path to html converted from ejs
                    if (Path.GetExtension(templatePath).ToLower() == ".html")
                    {
                        // If template path contains path to converted html template then delete it
                        File.Delete(templatePath);
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private static string ReplaceFileElementsWithMetaData(string templatePath, List<ContentMetaData> metaDataList, string outputFilePath)
        {
            string htmlContent = File.ReadAllText(templatePath);

            foreach (ContentMetaData metaData in metaDataList)
            {
                htmlContent = htmlContent.Replace($"{{{metaData.Placeholder}}}", metaData.Content);
            }

            string directoryPath = Path.GetDirectoryName(outputFilePath);
            string tempHtmlFilePath = Path.Combine(directoryPath, "Modified");
            string tempHtmlFile = Path.Combine(tempHtmlFilePath, "modifiedHtml.html");

            if (!Directory.Exists(tempHtmlFilePath))
            {
                Directory.CreateDirectory(tempHtmlFilePath);
            }

            File.WriteAllText(tempHtmlFile, htmlContent);
            return tempHtmlFile;
        }

        private static void ConvertHtmlToPdf(string toolFolderAbsolutePath, string modifiedHtmlFilePath, string outputFilePath)
        {
            string wkHtmlToPdfPath = "cmd.exe";
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                wkHtmlToPdfPath = "wkhtmltopdf";
            }

            /*
             * FIXME: Issue if tools file path has spaces in between
             */
            string arguments = HtmlToPdfArgumentsBasedOnOS(toolFolderAbsolutePath, modifiedHtmlFilePath, outputFilePath);

            ProcessStartInfo psi = new ProcessStartInfo
            {
                FileName = wkHtmlToPdfPath,
                Arguments = arguments,
                RedirectStandardOutput = false,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process())
            {
                process.StartInfo = psi;
                process.Start();
                process.WaitForExit();
                string output = process.StandardOutput.ReadToEnd();
                string errors = process.StandardError.ReadToEnd();
            }

            File.Delete(modifiedHtmlFilePath);
        }

        private static string ConvertEjsToHTML(string ejsFilePath, string outputFilePath, string ejsDataJson)
        {
            // Generate directory
            string directoryPath = Path.GetDirectoryName(outputFilePath);
            string tempDirectoryFilePath = Path.Combine(directoryPath, "Temp");

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
            File.WriteAllText(ejsDataJsonFilePath, ejsDataJson);

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
                process.WaitForExit();
                string output = process.StandardOutput.ReadToEnd();
                string errors = process.StandardError.ReadToEnd();
            }

            // Delete json data file
            File.Delete(ejsDataJsonFilePath);

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

        private static string HtmlToPdfArgumentsBasedOnOS(string toolFolderAbsolutePath, string modifiedHtmlFilePath, string outputFilePath)
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                return $"/C {toolFolderAbsolutePath} \"{modifiedHtmlFilePath}\" \"{outputFilePath}\"";
            }
            else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                return $"{modifiedHtmlFilePath} {outputFilePath}";
            }
            else
            {
                throw new Exception("Unknown operating system");
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
                return $"{ejsFilePath} -f {ejsDataJsonFilePath} -o {tempHtmlFilePath}";
            }
            else
            {
                throw new Exception("Unknown operating system");
            }
        }
    }
}
