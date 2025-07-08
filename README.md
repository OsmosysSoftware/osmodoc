# OsmoDoc

**OsmoDoc** is a powerful .NET library designed to generate PDF and Word documents dynamically using templates and structured data.

## Key Features

### Word Document Generation

* Replace placeholders in paragraphs and table cells with actual text.
* Handle multiple placeholders in a single paragraph or table cell.
* Populate entire tables using structured data.
* Replace placeholder images while maintaining the original image's position and size.

### PDF Document Generation

* Convert HTML templates to PDFs with dynamic placeholder substitution.
* Supports both plain HTML and EJS (Embedded JavaScript Templates).

### EJS to PDF Support

* Enables advanced templating logic with JavaScript-based syntax.
* JSON string data is passed to the EJS template at runtime.

---

## Development Setup

### Windows

* `wkhtmltopdf.exe` is already included in the repo at `OsmoDoc.API/wwwroot/Tool/wkhtmltopdf.exe`.
* Alternatively, download it from [https://wkhtmltopdf.org/downloads.html](https://wkhtmltopdf.org/downloads.html) and set the path manually.
* Install [Node.js](https://nodejs.org/) and npm.
* Install EJS globally via npm:

```bash
npm install -g ejs
```

### Linux

Install the following dependencies using your terminal:

```bash
sudo apt-get update && sudo apt-get install -y --no-install-recommends \
    wkhtmltopdf \
    nodejs \
    npm
sudo chmod 755 /usr/bin/wkhtmltopdf
npm install -g --only=prod ejs
```

### Common Steps (Windows & Linux)

* Clone the repo and navigate to the root folder.
* Create a `.env` file and copy values from `.env.example`.
* Set your environment-specific values.

---

## 📦 Docker-Based Setup (Cross-platform)

Set up the backend application in a Docker-based environment (Windows or Linux) to run and test APIs.

### Prerequisites

1. Install [Docker](https://docs.docker.com/engine/install/)
2. Install [Docker Compose](https://docs.docker.com/compose/install/)
3. (Optional) Install [Docker Extension for VS Code](https://marketplace.visualstudio.com/items?itemName=ms-azuretools.vscode-docker)

### Environment Setup

Before proceeding with Docker commands, create a `.env` file in the root directory and populate it with required variables. You can copy from the provided `.env.example` file.

> **Note:** All required values such as `SERVER_PORT`, `REDIS_PORT`, `REDIS_HOST`, `COMPOSE_PROJECT_NAME`, etc., must be set in the `.env` file.

### Docker Commands

Once `.env` is set, execute the following commands:

```bash
docker compose build
docker compose up -d
```

The application will be accessible at `http://localhost:<SERVER_PORT>` as configured in your `.env`.

* Swagger UI (Development): `http://localhost:<SERVER_PORT>/swagger/index.html`
* API Access: `http://localhost:<SERVER_PORT>/<API>`

### 🛠 Troubleshooting

If build fails due to network issues:

```bash
docker system prune -a
docker compose build
docker compose up -d
```

Refer to [prune docs](https://docs.docker.com/config/pruning/) before using.

---

## Usage Guide

Sample usage for PDF (HTML + EJS) and Word generation is available in [`usage_guide.md`](docs/guides/usage_guide.md).

---

## Target Framework

* .NET 8.0

---

## Citations

* [OpenXML SDK](https://github.com/dotnet/Open-XML-SDK)
* [wkhtmltopdf](https://wkhtmltopdf.org/)

---

## License

This project is licensed under the [MIT License](https://github.com/OsmosysSoftware/osmodoc/blob/main/LICENSE).

---

## Acknowledgements

Thanks to all the contributors who helped improve OsmoDoc!

<a href="https://github.com/OsmosysSoftware/osmodoc/graphs/contributors">
  <img src="https://contrib.rocks/image?repo=OsmosysSoftware/osmodoc" alt="Contributors" />
</a>
