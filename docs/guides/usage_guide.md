# OsmoDoc Usage Guide

This guide provides an overview of how to use the OsmoDoc library to generate PDF and Word documents from various templates and data sources.

## Getting Started

OsmoDoc leverages `wkhtmltopdf` for PDF generation, so ensure it's installed and accessible on your system.

### Prerequisites

* **.NET Core SDK:** Ensure you have the .NET Core SDK installed on your machine.
* **wkhtmltopdf:**
    * **Windows:** Download and install `wkhtmltopdf` from their official website. You'll need to set the `OsmoDocPdfConfig.WkhtmltopdfPath` property to the executable's location.
    * **Linux:** Install `wkhtmltopdf` via your distribution's package manager (e.g., `sudo apt-get install wkhtmltopdf` for Debian/Ubuntu). No explicit path setting is needed if it's in your system's PATH.

### Installation

To integrate OsmoDoc into your project, add the necessary NuGet packages:

```bash
dotnet add package OsmoDoc
```

---

## PDF Document Generation

OsmoDoc can generate PDFs from HTML and EJS (Embedded JavaScript) templates.

### HTML to PDF
This example demonstrates how to generate a PDF from a simple HTML template with dynamic content replacement.

Example Code
```csharp
// Set the wkhtmltopdf path if on Windows (on Linux, wkhtmltopdf needs to be installed)
OsmoDocPdfConfig.WkhtmltopdfPath = @"<path-to-tool>\wkhtmltopdf.exe"; // Adjust path as needed

// Define template and output paths
string templatePath = @"<input-path>/html_template.html";
string outputFilePath = @"<output-path>/html_output.pdf";

// Create dummy data for placeholders
List<ContentMetaData> metaData = new List<ContentMetaData>
{
   new ContentMetaData { Placeholder = "CustomerName", Content = "Jatin Gupta" },
   new ContentMetaData { Placeholder = "InvoiceNumber", Content = "GE202201484" },
   new ContentMetaData { Placeholder = "Amount", Content = "10000" }
};

try
{
   Console.WriteLine("Generating PDF...");
   await PdfDocumentGenerator.GeneratePdf(templatePath, metaData, outputFilePath, false, "");
   Console.WriteLine($"PDF generated successfully at {outputFilePath}");
}
catch (Exception ex)
{
   Console.WriteLine($"Error generating PDF: {ex.Message}");
}
```

### Explanation
- **OsmoDocPdfConfig.WkhtmltopdfPath**: Crucial for Windows users. Set this to the absolute path of your wkhtmltopdf.exe executable.
- **templatePath**: The path to your HTML template file.
- **metaData**: A List<ContentMetaData> where each ContentMetaData object defines a Placeholder (e.g., {{CustomerName}} in your HTML) and its corresponding Content.
- **outputFilePath**: The desired path for the generated PDF file.
- **PdfDocumentGenerator**.GeneratePdf: The core method for PDF generation.
    - **templatePath**: Path to the HTML template.
    - **metaData**: The list of placeholder data.
    - **outputFilePath**: Where to save the generated PDF.
    - **false**: Indicates that this is not an EJS template.
    - **""**: An empty string for EJS data (not applicable here).


### EJS to PDF
OsmoDoc also supports generating PDFs from EJS templates, allowing for more complex templating logic using JavaScript.

Example Code
```csharp
// Set the wkhtmltopdf path if on Windows (on Linux, wkhtmltopdf needs to be installed)
OsmoDocPdfConfig.WkhtmltopdfPath = @"<path-to-tool>\wkhtmltopdf.exe"; // Adjust path as needed

string ejsTemplatePath = "<input-path>/ejs_template.ejs";
string ejsOutputFilePath = "<output-path>/ejs_output.pdf";
string ejsDataJson = "{\"title\": \"EJS Test\", \"user\": {\"name\": \"Jane\"}}";

try
{
   Console.WriteLine("Generating EJS PDF...");
   await PdfDocumentGenerator.GeneratePdf(ejsTemplatePath, new List<ContentMetaData>(), ejsOutputFilePath, true, ejsDataJson);
   Console.WriteLine($"EJS PDF generated successfully at {ejsOutputFilePath}");
}
catch (Exception ex)
{
   Console.WriteLine($"Error generating EJS PDF: {ex.Message}");
}
```

### Explanation
- **OsmoDocPdfConfig.WkhtmltopdfPath**: Crucial for Windows users. Set this to the absolute path of your wkhtmltopdf.exe executable.
- **ejsTemplatePath**: The path to your EJS template file.
- **ejsDataJson**: A JSON string containing the data that will be passed to the EJS template. EJS templates use <%= variableName %> syntax to access this data.
- **PdfDocumentGenerator**.GeneratePdf:
    - **ejsTemplatePath**: Path to the EJS template.
    - **new List<ContentMetaData>()**: Placeholder list is empty as EJS handles data differently.
    - **ejsOutputFilePath**: Where to save the generated PDF.
    - **true**: Crucially, this indicates that it is an EJS template.
    - **ejsDataJson**: The JSON data string for the EJS template.

---

## Word Document Generation
OsmoDoc can populate Word .docx templates with text, table data, and images.

Example Code
```csharp
string wordTemplatePath = @"<input-path>/word_template.docx";
string wordOutputFilePath = @"<output-path>/word_output.docx";

// Text placeholders (optional)
List<ContentData> placeholders = new List<ContentData>()
{
    new ContentData
    {
        Placeholder = "InvoiceNo",
        Content = "INV-20250618",
        ContentType = ContentType.Text,
        ParentBody = ParentBody.None
    },
    new ContentData
    {
        Placeholder = "InvoiceDate",
        Content = "18 June 2025",
        ContentType = ContentType.Text,
        ParentBody = ParentBody.None
    },
    new ContentData
    {
        Placeholder = "TableCellNote", // Placeholder within a table cell
        Content = "Thanks for using OsmoDoc",
        ContentType = ContentType.Text,
        ParentBody = ParentBody.Table
    },
    new ContentData
    {
        Placeholder = "CustomerName",
        Content = "John Doe",
        ContentType = ContentType.Text,
        ParentBody = ParentBody.None
    }
};

// Table data example
List<TableData> tablesData = new List<TableData>()
{
    new TableData()
    {
        TablePos = 1, // The 1-based index of the table in the Word document
        Data = new List<Dictionary<string, string>>()
        {
            new Dictionary<string, string>()
            {
                { "Item", "Laptop" },
                { "Qty", "2" },
                { "Price", "60000" }
            },
            new Dictionary<string, string>()
            {
                { "Item", "Mouse" },
                { "Qty", "5" },
                { "Price", "500" }
            }
        }
    },
    new TableData()
    {
        TablePos = 2, // The 1-based index of the second table
        Data = new List<Dictionary<string, string>>()
        {
            new Dictionary<string, string>()
            {
                { "TaxType", "CGST" },
                { "Amount", "900" }
            },
            new Dictionary<string, string>()
            {
                { "TaxType", "SGST" },
                { "Amount", "600" }
            }
        }
    }
};

// Image data for different source types
List<ImageData> images = new List<ImageData>()
{
    // Local file
    new ImageData
    {
        PlaceholderName = "Picture 2",  // Alt text of image placeholder in Word
        SourceType = ImageSourceType.LocalFile,
        Data = @"/home/jatingupta/Downloads/shanks.jpg"
    },

    // URL
    new ImageData
    {
        PlaceholderName = "Picture 1",
        SourceType = ImageSourceType.Url,
        Data = "https://www.sample-videos.com/img/Sample-jpg-image-500kb.jpg"
    },

    // Base64
    new ImageData
    {
        PlaceholderName = "Picture 3",
        SourceType = ImageSourceType.Base64,
        Data = Convert.ToBase64String(System.IO.File.ReadAllBytes(@"/home/jatingupta/Downloads/shanks.jpg")),
        ImageExtension = ".jpg"
    }
};

// Combine all data into a DocumentData object
DocumentData documentData = new DocumentData()
{
    Placeholders = placeholders,
    TablesData = tablesData,
    Images = images
};

if (!File.Exists(wordTemplatePath))
{
    Console.WriteLine($"No file exists at: {wordTemplatePath}");
}

try
{
    Console.WriteLine("Generating Word Document...");
    await WordDocumentGenerator.GenerateDocumentByTemplate(wordTemplatePath, documentData, wordOutputFilePath);
    Console.WriteLine($"Word Document generated successfully at {wordOutputFilePath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error generating Word Document: {ex.Message}");
}
```

Explanation
- **wordTemplatePath**: The path to your .docx template file.
- **placeholders**: A List<ContentData> for replacing text placeholders in your Word document.
    - **Placeholder**: The exact text placeholder in your Word document (e.g., {{InvoiceNo}} or simply InvoiceNo if it's plain text).
    - **Content**: The text to replace the placeholder with.
    - **ContentType**: Specifies the type of content (e.g., ContentType.Text).
    - **ParentBody**: Indicates if the placeholder is within a regular document body or a table. Use ParentBody.Table for placeholders inside tables.
- **tablesData**: A List<TableData> to populate tables within your Word document.
    - **TablePos**: The 1-based index of the table in your Word document. The first table is 1, the second is 2, and so on.
    - **Data**: A List<Dictionary<string, string>> where each dictionary represents a row. The keys of the dictionary should match the column headers/placeholders in your template table (e.g., "Item", "Qty", "Price").
- **images**: A List<ImageData> to insert images into your Word document.
    - **PlaceholderName**: The alt text of the image placeholder in your Word document. You set this in Word by right-clicking an image placeholder, choosing "Format Picture" (or similar), and finding the "Alt Text" or "Description" field.
    - **SourceType**: Specifies the source of the image: LocalFile, Url, or Base64.
    - **Data**: The path to the local file, the URL, or the Base64 string of the image.
    - **ImageExtension**: Required for Base64 images to specify the file extension (e.g., .jpg, .png).
- **DocumentData**: This object aggregates all the placeholder, table, and image data for the document generation.
- **WordDocumentGenerator.GenerateDocumentByTemplate**: The method used to generate the Word document.
    - **wordTemplatePath**: Path to the Word template.
    - **documentData**: The aggregated data for the document.
    - **wordOutputFilePath**: Where to save the generated Word document.

--- 

### Note
- Templates used in the above examples are available [here](docs/templates).