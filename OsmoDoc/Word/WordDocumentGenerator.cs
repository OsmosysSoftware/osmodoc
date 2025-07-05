using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OsmoDoc.Word.Models;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using IOPath = System.IO.Path;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Net.Http;
using System.Diagnostics;

namespace OsmoDoc.Word;

/// <summary>
/// Provides functionality to generate Word documents based on templates and data.
/// </summary>
public static class WordDocumentGenerator
{
    private const string PlaceholderPattern = @"{[a-zA-Z]+}";

    /// <summary>
    /// Generates a Word document based on a template, replaces placeholders with data, and saves it to the specified output file path.
    /// </summary>
    /// <param name="templateFilePath">The file path of the template document.</param>
    /// <param name="documentData">The data to replace the placeholders in the template.</param>
    /// <param name="outputFilePath">The file path to save the generated document.</param>
    public async static Task GenerateDocumentByTemplate(string templateFilePath, DocumentData documentData, string outputFilePath)
    {
        if (string.IsNullOrWhiteSpace(templateFilePath))
        {
            throw new ArgumentNullException(nameof(templateFilePath));
        }

        if (documentData == null)
        {
            throw new ArgumentNullException(nameof(documentData));
        }

        if (string.IsNullOrWhiteSpace(outputFilePath))
        {
            throw new ArgumentNullException(nameof(outputFilePath));
        }

        List<ContentData> contentData = documentData.Placeholders;
        List<TableData> tablesData = documentData.TablesData;

        // Creating dictionaries for each type of placeholders
        Dictionary<string, string> textPlaceholders = new Dictionary<string, string>();
        Dictionary<string, string> tableContentPlaceholders = new Dictionary<string, string>();
        Dictionary<string, string> imagePlaceholders = new Dictionary<string, string>();

        foreach (ContentData content in contentData)
        {
            if (content.ParentBody == ParentBody.None && content.ContentType == ContentType.Text)
            {
                string placeholder = "{" + content.Placeholder + "}";
                textPlaceholders.TryAdd(placeholder, content.Content);
            }
            else if (content.ParentBody == ParentBody.Table && content.ContentType == ContentType.Text)
            {
                string placeholder = "{" + content.Placeholder + "}";
                tableContentPlaceholders.TryAdd(placeholder, content.Content);
            }
        }

        // Create document of the template
        XWPFDocument document = await GetXWPFDocument(templateFilePath);

        // For each element in the document
        foreach (IBodyElement element in document.BodyElements)
        {
            if (element.ElementType == BodyElementType.PARAGRAPH)
            {
                // If element is a paragraph
                XWPFParagraph paragraph = (XWPFParagraph)element;

                // If the paragraph is empty string or the placeholder regex does not match then continue
                if (paragraph.ParagraphText == string.Empty || !new Regex(PlaceholderPattern).IsMatch(paragraph.ParagraphText))
                {
                    continue;
                }

                // Replace placeholders in paragraph with values
                paragraph = ReplacePlaceholdersOnBody(paragraph, textPlaceholders);
            }
            else if (element.ElementType == BodyElementType.TABLE)
            {
                // If element is a table
                XWPFTable table = (XWPFTable)element;

                // Replace placeholders in a table
                table = ReplacePlaceholderOnTables(table, tableContentPlaceholders);

                // Populate the table with data if it is passed in tablesData list
                foreach (TableData insertData in tablesData)
                {
                    if (insertData.TablePos >= 1 && insertData.TablePos <= document.Tables.Count && table == document.Tables[insertData.TablePos - 1])
                    {
                        table = PopulateTable(table, insertData);
                    }
                }
            }
        }

        // Write the document to output file path and close the document
        WriteDocument(document, outputFilePath);
        document.Close();

        /*
            * Image Replacement is done after writing the document here,
            * because for Text Replacement, NPOI package is being used
            * and for Image Replacement, OpeXML package is used.
            * Since both the packages have different execution method, so they are handled separately
            */
        // Replace all the image placeholders in the output file
        await ProcessImagePlaceholders(outputFilePath, documentData.Images);
    }

    /// <summary>
    /// Retrieves an instance of XWPFDocument from the specified document file path.
    /// </summary>
    /// <param name="docFilePath">The file path of the Word document.</param>
    /// <returns>An instance of XWPFDocument representing the Word document.</returns>
    private async static Task<XWPFDocument> GetXWPFDocument(string docFilePath)
    {
        byte[] fileBytes = await File.ReadAllBytesAsync(docFilePath);
        using MemoryStream memoryStream = new MemoryStream(fileBytes);
        return new XWPFDocument(memoryStream);
    }

    /// <summary>
    /// Writes the XWPFDocument to the specified file path.
    /// </summary>
    /// <param name="document">The XWPFDocument to write.</param>
    /// <param name="filePath">The file path to save the document.</param>
    private static void WriteDocument(XWPFDocument document, string filePath)
    {
        string? directory = IOPath.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
        }

        using (FileStream writeStream = File.Create(filePath))
        {
            document.Write(writeStream);
        }
    }

    /// <summary>
    /// Replaces the text placeholders in a paragraph with the specified values.
    /// </summary>
    /// <param name="paragraph">The XWPFParagraph containing the placeholders.</param>
    /// <param name="textPlaceholders">The dictionary of text placeholders and their corresponding values.</param>
    /// <returns>The updated XWPFParagraph.</returns>
    private static XWPFParagraph ReplacePlaceholdersOnBody(XWPFParagraph paragraph, Dictionary<string, string> textPlaceholders)
    {
        // Get a list of all placeholders in the current paragraph
        List<string> placeholdersTobeReplaced = Regex.Matches(paragraph.ParagraphText, PlaceholderPattern)
                                                                 .Cast<Match>()
                                                                 .Select(s => s.Groups[0].Value).ToList();

        // For each placeholder in paragraph
        foreach (string placeholder in placeholdersTobeReplaced)
        {
            // Replace text placeholders in paragraph with values
            if (textPlaceholders.ContainsKey(placeholder))
            {
                paragraph.ReplaceText(placeholder, textPlaceholders[placeholder]);
            }

            paragraph.SpacingAfter = 0;
        }

        return paragraph;
    }

    /// <summary>
    /// Replaces the text placeholders in a table with the specified values.
    /// </summary>
    /// <param name="table">The XWPFTable containing the placeholders.</param>
    /// <param name="tableContentPlaceholders">The dictionary of table content placeholders and their corresponding values.</param>
    /// <returns>The updated XWPFTable.</returns>
    private static XWPFTable ReplacePlaceholderOnTables(XWPFTable table, Dictionary<string, string> tableContentPlaceholders)
    {
        // Loop through each cell of the table
        foreach (XWPFTableRow row in table.Rows)
        {
            foreach (XWPFTableCell cell in row.GetTableCells())
            {
                foreach (XWPFParagraph paragraph in cell.Paragraphs)
                {
                    // Get a list of all placeholders in the current cell
                    List<string> placeholdersTobeReplaced = Regex.Matches(paragraph.ParagraphText, PlaceholderPattern)
                                                             .Cast<Match>()
                                                             .Select(s => s.Groups[0].Value).ToList();

                    // For each placeholder in the cell
                    foreach (string placeholder in placeholdersTobeReplaced)
                    {
                        // replace the placeholder with its value
                        if (tableContentPlaceholders.ContainsKey(placeholder))
                        {
                            paragraph.ReplaceText(placeholder, tableContentPlaceholders[placeholder]);
                        }
                    }
                }
            }
        }

        return table;
    }

    /// <summary>
    /// Populates a table with the specified data.
    /// </summary>
    /// <param name="table">The XWPFTable to populate.</param>
    /// <param name="tableData">The data to populate the table.</param>
    /// <returns>The updated XWPFTable.</returns>
    private static XWPFTable PopulateTable(XWPFTable table, TableData tableData)
    {
        // Get the header row
        XWPFTableRow headerRow = table.GetRow(0);

        // Return if no header row found or if it does not have any column
        if (headerRow == null || headerRow.GetTableCells() == null || headerRow.GetTableCells().Count <= 0)
        {
            return table;
        }

        // For each row's data stored in table data
        foreach (Dictionary<string, string> rowData in tableData.Data)
        {
            XWPFTableRow row = table.CreateRow(); // This is a DATA row, not header

            int columnCount = headerRow.GetTableCells().Count; // Read from header
            for (int cellNumber = 0; cellNumber < columnCount; cellNumber++)
            {
                // Ensure THIS data row has enough cells
                while (row.GetTableCells().Count <= cellNumber)
                {
                    row.AddNewTableCell();
                }

                // Now populate the cell in this data row
                XWPFTableCell cell = row.GetCell(cellNumber);
                string columnHeader = headerRow.GetCell(cellNumber).GetText();
                if (rowData.ContainsKey(columnHeader))
                {
                    cell.SetText(rowData[columnHeader]);
                }
            }
        }

        return table;
    }

    /// <summary>
    /// Replaces the image placeholders in the output file with the specified images.
    /// </summary>
    /// <param name="documentPath">The ile path where the updated document will be saved.</param>
    /// <param name="images">The data structure for holding the images details.</param>

    private static async Task ProcessImagePlaceholders(
        string documentPath,
        List<ImageData> images)
    {
        if (images == null || !images.Any())
        {
            return;
        }

        List<string> tempFiles = new List<string>();

        try
        {
            byte[] docBytes = await File.ReadAllBytesAsync(documentPath);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                await memoryStream.WriteAsync(docBytes);
                memoryStream.Position = 0;

                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(memoryStream, true))
                {
                    MainDocumentPart? mainPart = wordDocument.MainDocumentPart;
                    if (mainPart == null)
                    {
                        return;
                    }

                    List<Drawing> drawings = mainPart.Document.Descendants<Drawing>().ToList();

                    foreach (ImageData img in images)
                    {
                        try
                        {
                            string tempFilePath = await PrepareImageFile(img);
                            tempFiles.Add(tempFilePath);

                            Drawing? drawing = drawings.FirstOrDefault(d =>
                                d.Descendants<DocProperties>()
                                 .Any(dp => dp.Name == img.PlaceholderName));

                            if (drawing == null)
                            {
                                continue;
                            }

                            foreach (Blip blip in drawing.Descendants<Blip>())
                            {
                                if (blip.Embed?.Value == null)
                                {
                                    continue;
                                }

                                OpenXmlPart imagePart = mainPart.GetPartById(blip.Embed!);
                                using (Stream partStream = imagePart.GetStream(FileMode.Create))
                                {
                                    await using FileStream fileStream = File.OpenRead(tempFilePath);
                                    await fileStream.CopyToAsync(partStream);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Log error but continue with other images
                            Debug.WriteLine($"Failed to process image {img.PlaceholderName}: {ex.Message}");
                        }
                    }
                }

                // Save the modified document
                memoryStream.Position = 0;
                using (FileStream fileStream = new FileStream(documentPath, FileMode.Create))
                {
                    await memoryStream.CopyToAsync(fileStream);
                }
            }
        }
        finally
        {
            // Clean up temp files
            foreach (string file in tempFiles)
            {
                try { File.Delete(file); }
                catch { /* Ignore cleanup errors */ }
            }
        }
    }

    private static async Task<string> PrepareImageFile(ImageData imageData)
    {
        string tempFilePath = System.IO.Path.GetTempFileName();

        if (!string.IsNullOrEmpty(imageData.ImageExtension))
        {
            tempFilePath = System.IO.Path.ChangeExtension(tempFilePath, imageData.ImageExtension);
        }

        switch (imageData.SourceType)
        {
            case ImageSourceType.Base64:
                await File.WriteAllBytesAsync(
                    tempFilePath,
                    Convert.FromBase64String(imageData.Data));
                break;

            case ImageSourceType.LocalFile:
                if (!File.Exists(imageData.Data))
                {
                    throw new FileNotFoundException("Image file not found", imageData.Data);
                }

                File.Copy(imageData.Data, tempFilePath, true);
                break;

            case ImageSourceType.Url:
                using (HttpClient httpClient = new HttpClient())
                {
                    byte[] bytes = await httpClient.GetByteArrayAsync(imageData.Data);
                    await File.WriteAllBytesAsync(tempFilePath, bytes);
                }
                break;

            default:
                throw new ArgumentOutOfRangeException();
        }

        return tempFilePath;
    }
}