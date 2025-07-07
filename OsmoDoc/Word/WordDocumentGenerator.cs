using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Logging.Abstractions;
using OsmoDoc.Word.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using IOPath = System.IO.Path;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableCellProperties = DocumentFormat.OpenXml.Wordprocessing.TableCellProperties;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace OsmoDoc.Word;

/// <summary>
/// Provides functionality to generate Word documents based on templates and data.
/// </summary>
public static class WordDocumentGenerator
{
    private const string PlaceholderPattern = @"{[a-zA-Z][a-zA-Z0-9_-]*}";
    private static ILogger _logger = NullLogger.Instance;

    /// <summary>
    /// Configures logging for the WordDocumentGenerator
    /// </summary>
    /// <param name="logger">Logger instance to use</param>
    public static void ConfigureLogging(ILogger logger)
    {
        _logger = logger ?? NullLogger.Instance;
    }

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

        try
        {
            // Copy template to output location
            File.Copy(templateFilePath, outputFilePath, true);

            using (WordprocessingDocument document = WordprocessingDocument.Open(outputFilePath, true))
            {
                if (document.MainDocumentPart == null)
                {
                    throw new InvalidOperationException("Document does not contain a main document part.");
                }

                // Create dictionaries for each type of placeholders
                Dictionary<string, string> textPlaceholders = documentData.Placeholders
                    .Where(content => content.ParentBody == ParentBody.None && content.ContentType == ContentType.Text)
                    .ToDictionary(content => "{" + content.Placeholder + "}", content => content.Content);

                Dictionary<string, string> tableContentPlaceholders = documentData.Placeholders
                    .Where(content => content.ParentBody == ParentBody.Table && content.ContentType == ContentType.Text)
                    .ToDictionary(content => "{" + content.Placeholder + "}", content => content.Content);

                // Replace text placeholders in main document
                ReplaceTextPlaceholders(document.MainDocumentPart.Document, textPlaceholders);

                // Replace table placeholders and populate tables
                ProcessTables(document.MainDocumentPart.Document, tableContentPlaceholders, documentData.TablesData);

                // Process images
                await ProcessImagePlaceholders(document, documentData.Images);

                // Save the document
                document.Save();
            }
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, $"Failed to generate document from template: {templateFilePath}");
            throw; // Re-throw for consumers to handle
        }
    }

    /// <summary>
    /// Replaces text placeholders in the document.
    /// </summary>
    /// <param name="document">The document to process.</param>
    /// <param name="textPlaceholders">Dictionary of placeholders and their replacement values.</param>
    private static void ReplaceTextPlaceholders(Document document, Dictionary<string, string> textPlaceholders)
    {
        if (textPlaceholders.Count == 0)
        {
            return;
        }

        // Process all paragraphs in the document
        List<Paragraph> paragraphs = document.Descendants<Paragraph>().ToList();

        foreach (Paragraph paragraph in paragraphs)
        {
            // Get paragraph text to check for placeholders
            string paragraphText = GetParagraphText(paragraph);

            if (string.IsNullOrEmpty(paragraphText) || !Regex.IsMatch(paragraphText, PlaceholderPattern))
            {
                continue;
            }

            // Replace placeholders in this paragraph
            ReplacePlaceholdersInParagraph(paragraph, textPlaceholders);
        }
    }

    /// <summary>
    /// Gets the text content of a paragraph.
    /// </summary>
    /// <param name="paragraph">The paragraph to get text from.</param>
    /// <returns>The text content of the paragraph.</returns>
    private static string GetParagraphText(Paragraph paragraph)
    {
        return string.Join("", paragraph.Descendants<Text>().Select(t => t.Text));
    }

    /// <summary>
    /// Replaces placeholders in a specific paragraph.
    /// </summary>
    /// <param name="paragraph">The paragraph to process.</param>
    /// <param name="placeholders">Dictionary of placeholders and their replacement values.</param>
    private static void ReplacePlaceholdersInParagraph(Paragraph paragraph, Dictionary<string, string> placeholders)
    {
        if (placeholders.Count == 0)
        {
            return;
        }

        // Get all text elements from the paragraph
        List<Text> textElements = paragraph.Descendants<Text>().ToList();
        if (textElements.Count == 0)
        {
            return;
        }

        // Concatenate all text to get the full paragraph content
        string fullText = string.Join("", textElements.Select(t => t.Text));

        // Check if any placeholders exist in the full text
        bool hasPlaceholders = placeholders.Keys.Any(placeholder => fullText.Contains(placeholder));
        if (!hasPlaceholders)
        {
            return;
        }

        // Replace placeholders in the full text
        string replacedText = fullText;
        foreach (KeyValuePair<string, string> placeholder in placeholders)
        {
            replacedText = replacedText.Replace(placeholder.Key, placeholder.Value);
        }

        // If no changes were made, return
        if (replacedText == fullText)
        {
            return;
        }

        // Clear existing text elements and create a single new one
        // This preserves the paragraph structure while ensuring text continuity
        foreach (Text textElement in textElements)
        {
            textElement.Text = "";
        }

        // Use the first text element to hold all the replaced content
        if (textElements.Count > 0)
        {
            textElements[0].Text = replacedText;
        }
    }

    /// <summary>
    /// Processes tables for placeholder replacement and data population.
    /// </summary>
    /// <param name="document">The document containing tables.</param>
    /// <param name="tableContentPlaceholders">Dictionary of table placeholders and their replacement values.</param>
    /// <param name="tablesData">List of table data to populate.</param>
    private static void ProcessTables(Document document, Dictionary<string, string> tableContentPlaceholders, List<TableData> tablesData)
    {
        List<Table> tables = document.Descendants<Table>().ToList();

        foreach (Table table in tables)
        {
            // Replace placeholders in table cells
            ReplaceTablePlaceholders(table, tableContentPlaceholders);

            // Populate table with data if applicable
            int tableIndex = tables.IndexOf(table);
            TableData? tableData = tablesData.FirstOrDefault(td => td.TablePos == tableIndex + 1);

            if (tableData != null)
            {
                PopulateTable(table, tableData);
            }
        }
    }

    /// <summary>
    /// Replaces placeholders in table cells.
    /// </summary>
    /// <param name="table">The table to process.</param>
    /// <param name="tableContentPlaceholders">Dictionary of placeholders and their replacement values.</param>
    private static void ReplaceTablePlaceholders(Table table, Dictionary<string, string> tableContentPlaceholders)
    {
        if (tableContentPlaceholders.Count == 0)
        {
            return;
        }

        List<TableRow> tableRows = table.Elements<TableRow>().ToList();

        foreach (TableRow row in tableRows)
        {
            List<TableCell> cells = row.Elements<TableCell>().ToList();

            foreach (TableCell cell in cells)
            {
                List<Paragraph> paragraphs = cell.Elements<Paragraph>().ToList();

                foreach (Paragraph paragraph in paragraphs)
                {
                    string paragraphText = GetParagraphText(paragraph);

                    if (string.IsNullOrEmpty(paragraphText) || !Regex.IsMatch(paragraphText, PlaceholderPattern))
                    {
                        continue;
                    }

                    ReplacePlaceholdersInParagraph(paragraph, tableContentPlaceholders);
                }
            }
        }
    }

    /// <summary>
    /// Populates a table with data rows.
    /// </summary>
    /// <param name="table">The table to populate.</param>
    /// <param name="tableData">The data to populate the table with.</param>
    private static void PopulateTable(Table table, TableData tableData)
    {
        TableRow? headerRow = table.Elements<TableRow>().FirstOrDefault();
        if (headerRow == null)
        {
            return;
        }

        List<TableCell> headerCells = headerRow.Elements<TableCell>().ToList();
        if (headerCells.Count == 0)
        {
            return;
        }

        // Get column headers
        List<string> columnHeaders = headerCells.Select(cell =>
        {
            Paragraph? firstParagraph = cell.Elements<Paragraph>().FirstOrDefault();
            if (firstParagraph != null)
            {
                return string.Join("", firstParagraph.Descendants<Text>().Select(t => t.Text));
            }
            return "";
        }).ToList();

        // Add data rows
        foreach (Dictionary<string, string> rowData in tableData.Data)
        {
            TableRow newRow = new TableRow();

            for (int i = 0; i < columnHeaders.Count; i++)
            {
                string cellValue = rowData.ContainsKey(columnHeaders[i]) ? rowData[columnHeaders[i]] : "";

                TableCell cell = new TableCell(
                    new Paragraph(
                        new Run(
                            new Text(cellValue))));

                // Copy formatting from header cell if available
                if (i < headerCells.Count)
                {
                    TableCellProperties? headerProps = headerCells[i].TableCellProperties;
                    if (headerProps != null)
                    {
                        cell.TableCellProperties = (TableCellProperties)headerProps.CloneNode(true);
                    }
                }

                newRow.Append(cell);
            }

            table.Append(newRow);
        }
    }

    /// <summary>
    /// Processes image placeholders in the document.
    /// </summary>
    /// <param name="document">The Word document.</param>
    /// <param name="images">List of image data to process.</param>
    private static async Task ProcessImagePlaceholders(WordprocessingDocument document, List<ImageData> images)
    {
        if (images == null || !images.Any())
        {
            return;
        }

        List<string> tempFiles = new List<string>();

        try
        {
            MainDocumentPart? mainPart = document.MainDocumentPart;
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
                         .Any(dp => dp.Description == img.PlaceholderName));

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
                        using (FileStream fileStream = File.OpenRead(tempFilePath))
                        {
                            await fileStream.CopyToAsync(partStream);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"Failed to process image placeholder: {img.PlaceholderName}");
                }
            }
        }
        finally
        {
            // Clean up temp files
            foreach (string file in tempFiles)
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, $"Failed to delete temporary file: {file}.");
                }
            }
        }
    }

    /// <summary>
    /// Prepares an image file from various sources (Base64, local file, URL).
    /// </summary>
    /// <param name="imageData">The image data containing source information.</param>
    /// <returns>Path to the prepared temporary image file.</returns>
    private static async Task<string> PrepareImageFile(ImageData imageData)
    {
        string tempFilePath = IOPath.GetTempFileName();

        if (!string.IsNullOrEmpty(imageData.ImageExtension))
        {
            // Define allowed image extensions
            string[] allowedExtensions = new[] { ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".svg" };
            string extension = imageData.ImageExtension.StartsWith(".")
                ? imageData.ImageExtension
                : "." + imageData.ImageExtension;

            if (!allowedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException($"Invalid image extension: {imageData.ImageExtension}");
            }

            tempFilePath = IOPath.ChangeExtension(tempFilePath, extension);
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