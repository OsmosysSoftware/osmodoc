# OsmoDoc
OsmoDoc is a library with the following functions
1. **Generate Word documents** - Read Word document files as a template and replace the placeholder with actual data.
2. **Generate PDF documents** - Read an HTML file as a template and replace placeholders with actual data. Convert the HTML file to PDF

# Features

## Word document generation
- Replace placeholders in text paragraph with values.
- Replace placeholders in tables.
- Multiple placeholders in the same table cell/line/paragraph can be replaced.
- Populate table with new data.
- Replace images with image's placeholder. The image's position will be maintained based on the position of its placeholder image. Image size will also be maintained based on placeholder image.

## PDF document generation
- Converts an HTML document to PDF.
- Replace placeholders in the document with actual string data.

# How to set up the Application in a Docker-based environment (Linux)

Setting up the app in a Docker-based environment enables developers of non-Windows origins to run the backend application on their machine to test the APIs.

## Steps

1. [Install Docker](https://docs.docker.com/engine/install/) on your machine. Choose to follow the instructions based on your device OS.
2. [Install Docker Compose](https://docs.docker.com/compose/install/). A separate installation is required for Linux-based OS. If you are using Windows or macOS, installing the Docker Desktop app includes Docker Compose.
3. Clone the project `osmodoc`.
4. (Optional) [Install Docker Extension for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-docker).
5. In the root directory of the project, create a new file `.env`.
6. Copy data from [example template](.env.example) into `.env`. Then set suitable JWT key.
7. Set `environment` variables `ASPNETCORE_ENVIRONMENT` and `BUILD_CONFIGURATION` as per requirement in [docker-compose.yaml](./docker-compose.yaml). Ensure correct formatting:

#### Development
```yaml
      - BUILD_CONFIGURATION=Debug
      - ASPNETCORE_ENVIRONMENT=Development
```

#### Testing/Staging
```yaml
      - BUILD_CONFIGURATION=Release
      - ASPNETCORE_ENVIRONMENT=Development
```

#### Production
```yaml
      - BUILD_CONFIGURATION=Release
      - ASPNETCORE_ENVIRONMENT=Production
```

8. Ensure Docker is running.
9. Execute the following commands to dockerize `osmodoc` using `docker-compose.yaml`:

```shell
# build the container
docker compose build

# run the application
docker compose up
```

10. The project will run on `http://localhost:5000`. Please check [Troubleshooting](#troubleshooting) if the build failed.
11. You can access the **Swagger UI** at `http://localhost:5000/swagger/index.html` in **Development** Environment.
12. Test the API via **Postman**. The app can be accessed using `http://localhost:5000/<API>`.

## Troubleshooting

A known issue while building the container is the following:

```shell
E: failed to solve: process "/bin/sh -c <sample Dockerfile step>" did not complete successfully: exit code: 100
```

This is a network related issue where it is failing to fetch files from an external source. It can be verified in the **Docker logs**:

```shell
E: Failed to fetch http://sample/link/for.file Unable to connect to sample.download.location:80: [IP: ...]
```

**Solution:** Prune the failed build and rebuild the application using the following commands:

```shell
# prune all unused containers, networks, images, build cache
docker system prune -a

# rebuild the container
docker compose build

# run the application
docker compose up
```

**NOTE:** Please go through the [official documentation on prune command](https://docs.docker.com/config/pruning/) before using it.

# How to set up the library (Windows)

## Steps for installing wkhtmltopdf
- Go to website: https://wkhtmltopdf.org/downloads.html
- Select the version of wkhtmltopdf installer that you need to download based on your system requirements.
- Finish Installation.

## Including wkhtmltopdf executable file to build package
- Go to the location to the bin files of your project where the OsmoDoc DLL is located.
- Create a folder called Tools and place the wkhtmltopdf.exe file there. wkhtmltopdf.exe can be found in the Program Files in C directory after it is installed.

Note: We use a Temp folder to temporarily hold the modified HTML file before converting it to a PDF file. After the conversion is done, the temporary file is removed. The code is already provided with the location of the temp file, so no modification is required in the code, and the temp folder will be used automatically.

# Basic usage

## PDF generation

#### HTML TO PDF
```csharp
string htmlTemplateFilePath = @"C:\Path\To\Template.html";
string outputFilePath = @"C:\Path\To\GeneratedOutput.pdf";

List<ContentMetaData> contentList = new List<ContentMetaData>
{
    new ContentMetaData { Placeholder = "Incident UID", Content = "I-20230822-001" },
    new ContentMetaData { Placeholder = "Description", Content = "Suspicious activity reported" },
    new ContentMetaData { Placeholder = "Site", Content = "Headquarters" }
};

await PdfDocumentGenerator.GeneratePdf(htmlTemplateFilePath, contentList, outputFilePath, isEjsTemplate: false, serializedEjsDataJson: null);
```

#### EJS TO PDF
```csharp
string htmlTemplateFilePath = @"C:\Path\To\Template.ejs";
string outputFilePath = @"C:\Path\To\GeneratedOutput.pdf";
string serializedEjsDataJson = "{\"title\": \"EJS Test\", \"user\": {\"name\": \"Jane\"}}"

List<ContentMetaData> contentList = new List<ContentMetaData>{};

await PdfDocumentGenerator.GeneratePdf(htmlTemplateFilePath, contentList, outputFilePath, isEjsTemplate: true, serializedEjsDataJson: serializedEjsDataJson);
```

## Word document generation
```csharp
string templateFilePath = @"C:\Path\To\Template.docx";
string outputFilePath = @"C:\Path\To\GeneratedOutput.docx";

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
        Placeholder = "TableCellNote",
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
        TablePos = 1,
        Data = new List<Dictionary<string, string>>()
        {
            new Dictionary<string, string>()
            {
                { "Item", "Laptop" },
                { "Oty", "2" },
                { "Price", "60000" }
            },
            new Dictionary<string, string>()
            {
                { "Item", "Mouse" },
                { "Oty", "5" },
                { "Price", "500" }
            }
        }
    },
    new TableData()
    {
        TablePos = 2,
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
        PlaceholderName = "Picture 1",  // Alt text of image placeholder in Word
        SourceType = ImageSourceType.LocalFile,
        Data = @"C:\Images\logo.png"
    },

    // URL
    new ImageData
    {
        PlaceholderName = "Picture 2",
        SourceType = ImageSourceType.Url,
        Data = "https://example.com/image.jpg"
    },

    // Base64 (ImageExtension is required when SourceType is Base64)
    new ImageData
    {
        PlaceholderName = "Picture 3",
        SourceType = ImageSourceType.Base64,
        Data = "<base64-encoded-string>",
        ImageExtension = ".jpg"
    }
};

// Combine all document parts
DocumentData documentData = new DocumentData
{
    Placeholders = placeholders,
    TablesData = tablesData,
    Images = images
};

// Generate final Word document
await WordDocumentGenerator.GenerateDocumentByTemplate(templateFilePath, documentData, outputFilePath);
```

# Targeted frameworks
1. .NET Framework 8.0

# Citations
- [NPOI](https://github.com/nissl-lab/npoi)
- [OpenXML](https://github.com/dotnet/Open-XML-SDK)
- [wkhtmltopdf](https://wkhtmltopdf.org/)

# License
The OsmoDoc is licensed under the [MIT](https://github.com/OsmosysSoftware/osmodoc/blob/main/LICENSE) license.

## 👏 Big Thanks to Our Contributors

<a href="https://github.com/OsmosysSoftware/osmodoc/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=OsmosysSoftware/osmodoc" alt="Contributors" />
</a>

We appreciate the time and effort put in by all contributors to make this project better!