using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using OsmoDoc.Pptx.Models;
using OsmoDoc.Pptx.Services;

namespace OsmoDoc.Pptx;

public class PptxGenerator
{
    private readonly LlmSlideExtractorService _extractor;
    private readonly HtmlGeneratorService _htmlGen;
    private readonly PptxService _pptx;
    private readonly IWebHostEnvironment _env;

    public PptxGenerator(
        LlmSlideExtractorService extractor,
        HtmlGeneratorService htmlGen,
        PptxService pptx,
        IWebHostEnvironment env)
    {
        this._extractor = extractor;
        this._htmlGen = htmlGen;
        this._pptx = pptx;
        this._env = env;
    }

    /// <summary>
    /// Accepts free-form text, extracts slides, generates HTML and PPTX.
    /// </summary>
    public async Task<string> GeneratePresentationFromTextAsync(string userText, CancellationToken ct = default)
    {
        // Create per-request folder
        string requestId = Guid.NewGuid().ToString("N").Substring(0, 8);
        string requestDir = Path.Combine(this._env.ContentRootPath, "Generated", requestId);
        Directory.CreateDirectory(requestDir);

        // Step 1: LLM extract slides JSON
        SlideExtractionResult slidesResult = await this._extractor.ExtractSlidesAsync(userText, ct);

        // Step 2: Convert to HTML
        string filledHtml = this._htmlGen.GenerateHtml(slidesResult);

        // Save HTML inside request folder
        string debugHtmlPath = Path.Combine(requestDir, "debug.html");
        await File.WriteAllTextAsync(debugHtmlPath, filledHtml, ct);

        // Step 3: Generate PPTX into same request folder
        string fileName = $"presentation-{DateTime.UtcNow:yyyyMMddHHmmss}.pptx";
        string pptxPath = this._pptx.GeneratePptxFromHtmlString(filledHtml, fileName, requestDir);
        return pptxPath;
    }

    /// <summary>
    /// Accepts free-form text plus images, extracts slides, generates HTML and PPTX.
    /// </summary>
    public async Task<string> GeneratePresentationFromTextAndImagesAsync(string userText, IEnumerable<IFormFile> images, CancellationToken ct = default)
    {
        // Persist images and build catalog mapping ID -> relative path
        List<IFormFile> imageList = images?.ToList() ?? new List<IFormFile>();
        List<(string Id, string FileName, string RelativePath)> catalog = new List<(string Id, string FileName, string RelativePath)>();
        string requestId = Guid.NewGuid().ToString("N").Substring(0, 8);
        string requestDir = Path.Combine(this._env.ContentRootPath, "Generated", requestId);
        string imagesDir = Path.Combine(requestDir, "images");
        Directory.CreateDirectory(imagesDir);

        int idx = 1;
        foreach (IFormFile? f in imageList)
        {
            if (idx > this._maxImages)
            {
                break;
            }

            if (f.Length == 0)
            {
                continue;
            }

            string safeName = Path.GetFileName(f.FileName);
            string id = $"img{idx++}";
            string fullPath = Path.Combine(imagesDir, safeName);
            using (FileStream fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                await f.CopyToAsync(fs, ct);
            }
            string rel = Path.Combine("Generated", requestId, "images", safeName).Replace("\\", "/");
            catalog.Add((id, safeName, rel));
        }

        // Extract slides (catalog only mode, extractor rebuilds catalog by filenames/IDs ordering)
        SlideExtractionResult slidesResult = await this._extractor.ExtractSlidesAsync(userText, imageList, ct);

        // Inject image tags into HTML using mapping
        string filledHtml = this._htmlGen.GenerateHtml(slidesResult, catalog.ToDictionary(c => c.Id, c => c.RelativePath));

        // Save HTML in request folder
        string debugHtmlPath = Path.Combine(requestDir, "debug.html");
        await File.WriteAllTextAsync(debugHtmlPath, filledHtml, ct);

        string fileName = $"presentation-{DateTime.UtcNow:yyyyMMddHHmmss}.pptx";
        string pptxPath = this._pptx.GeneratePptxFromHtmlString(filledHtml, fileName, requestDir);
        return pptxPath;
    }

    private int _maxImages = 5; // default

    public void ConfigureFrom(IConfiguration config)
    {
        string? maxImagesRaw = Environment.GetEnvironmentVariable("MAX_IMAGES");

        if (int.TryParse(maxImagesRaw, out int v) && v > 0)
        {
            this._maxImages = v;
        }
    }
}
