using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using DocumentService.Pdf.Models;
using System.Diagnostics;

namespace DocumentService.Pdf
{     
    public class PdfDocumentGenerator
    {
        public static void GeneratePdfByTemplate(string toolFolderAbsolutePath, string templatePath, List<ContentMetaData> metaDataList, string outputFilePath)
        {
            try 
            { 
                if(!File.Exists(templatePath))
                {
                    throw new Exception("The file path you provided is not Valid"); 
                }
                
                string modifiedHtmlFilePath =  ReplaceFileElementsWithMetaData(templatePath, metaDataList, outputFilePath);
                ConvertHtmlToPdf(toolFolderAbsolutePath, modifiedHtmlFilePath, outputFilePath);
            } 
            catch(Exception e)
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
            string tempHtmlFilePath = Path.Combine(directoryPath, "Temp");
            string tempHtmlFile = Path.Combine(tempHtmlFilePath, "modifiedHtml.html");

            if (!Directory.Exists(tempHtmlFilePath))
            {
                Directory.CreateDirectory(tempHtmlFilePath);
            }

            File.WriteAllText(tempHtmlFile, htmlContent);    
            return tempHtmlFile;
        }

        private static void ConvertHtmlToPdf(string wkhtmltopdfExePath, string htmlFilePath, string outputPdfPath)
        {
            string arguments = $"-q --enable-local-file-access \"{htmlFilePath}\" \"{outputPdfPath}\"";

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = wkhtmltopdfExePath,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process())
            {
                process.StartInfo = startInfo;
                process.Start();
                process.WaitForExit();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
            }

            // Clean up the temporary HTML file safely
            if (File.Exists(htmlFilePath))
            {
                File.Delete(htmlFilePath);
            }
        }
    }
}
