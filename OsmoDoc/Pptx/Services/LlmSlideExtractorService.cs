using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using OsmoDoc.Pptx.Models;

namespace OsmoDoc.Pptx.Services;

public class LlmSlideExtractorService
{
    private readonly HttpClient _http;
    private readonly string _openAiKey;
    private readonly string _model;
    private readonly int _maxImages;

    public LlmSlideExtractorService(HttpClient http, IConfiguration config)
    {
        this._http = http;

        string? apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");

        if (string.IsNullOrWhiteSpace(apiKey))
        {
            throw new InvalidOperationException("OpenAI API key not configured");
        }

        this._openAiKey = apiKey;

        this._model = Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-5-mini";

        this._http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", this._openAiKey);

        string? maxImagesRaw = Environment.GetEnvironmentVariable("MAX_IMAGES");

        if (!int.TryParse(maxImagesRaw, out this._maxImages) || this._maxImages <= 0)
        {
            this._maxImages = 5;
        }
    }

    public Task<SlideExtractionResult> ExtractSlidesAsync(string userInput, CancellationToken ct = default)
        => this.ExtractSlidesInternalAsync(userInput, null, ct);

    public async Task<SlideExtractionResult> ExtractSlidesAsync(string userInput, IEnumerable<IFormFile> imageFiles, CancellationToken ct = default)
    {
        // Vision mode: build catalog plus base64 payload so model can inspect images.
        List<(string Id, string Mime, string Base64, string FileName)> imagesPayload = new List<(string Id, string Mime, string Base64, string FileName)>();
        int index = 1;
        foreach (IFormFile file in imageFiles ?? Enumerable.Empty<IFormFile>())
        {
            if (imagesPayload.Count >= this._maxImages)
            {
                break;
            }

            if (file.Length == 0)
            {
                continue;
            }

            using MemoryStream ms = new MemoryStream();
            await file.CopyToAsync(ms, ct);
            byte[] bytes = ms.ToArray();
            string id = $"img{index++}";
            string mime = string.IsNullOrWhiteSpace(file.ContentType) ? "image/png" : file.ContentType;
            imagesPayload.Add((id, mime, Convert.ToBase64String(bytes), file.FileName));
        }
        return await this.ExtractSlidesInternalAsync(userInput, imagesPayload, ct);
    }

    public Task<SlideExtractionResult> ExtractSlidesAsync(string userInput, IEnumerable<(string MimeType, string Base64Data)> images, CancellationToken ct = default)
    {
        // Accept pre-encoded images with synthetic IDs in order
        List<(string Id, string Mime, string Base64, string FileName)> list = new List<(string Id, string Mime, string Base64, string FileName)>();
        int i = 1;
        foreach ((string MimeType, string Base64Data) im in images ?? Enumerable.Empty<(string MimeType, string Base64Data)>())
        {
            if (list.Count >= this._maxImages)
            {
                break;
            }

            list.Add(($"img{i++}", im.MimeType, im.Base64Data, $"image{i}.png"));
        }
        return this.ExtractSlidesInternalAsync(userInput, list, ct);
    }

    private async Task<SlideExtractionResult> ExtractSlidesInternalAsync(string userInput, List<(string Id, string Mime, string Base64, string FileName)>? images, CancellationToken ct)
    {
        if (images != null && images.Count > this._maxImages)
        {
            images = images.Take(this._maxImages).ToList();
        }

        // Build catalog text from images list
        string catalogText = string.Empty;
        if (images != null && images.Count > 0)
        {
            IEnumerable<string> entries = images.Select(c => $"{{ id: '{c.Id}', filename: '{c.FileName}' }}");
            catalogText = "Image catalog: [" + string.Join(", ", entries) + "]\nWhen a slide uses one or more images add an 'images': ['img1','img2'] property referencing ONLY these IDs.";
        }

        string systemPrompt = "You are an assistant that extracts slide data from user instructions and attached images.\n" +
                "Return ONLY valid JSON with this schema: { \"slides\": [ { \"layout\": \"title-slide|content-slide|two-content-slide|thank-you-slide\", \"title\": \"string\", \"subtitle\": \"string\", \"content\": \"string\", \"bullets\": [\"string\"], \"images\": [\"img1\", \"img2\"] } ] }\n" +
                "Rules:\n- first slide => title-slide\n- main slides => content-slide\n- slides that pair textual explanation with a referenced image => two-content-slide\n- last slide with a thank you message => thank-you-slide\n- bullets for lists, content for paragraphs\n- only use image IDs from the provided catalog (do not invent)\n- omit images field if none used\n- Output ONLY JSON.";

        object payload;
        if (images == null || images.Count == 0)
        {
            string userPromptCombined = (catalogText.Length > 0)
                ? $"User presentation description:\n{userInput}\n\n{catalogText}"
                : $"User presentation description:\n{userInput}";
            payload = new
            {
                model = this._model,
                messages = new object[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = userPromptCombined }
                }
            };
        }
        else
        {
            List<object> parts = new List<object>();
            string intro = $"User presentation description:\n{userInput}\n\n{catalogText}";
            parts.Add(new { type = "text", text = intro });
            foreach ((string Id, string Mime, string Base64, string FileName) im in images)
            {
                parts.Add(new { type = "text", text = $"Image id: {im.Id}" });
                parts.Add(new { type = "image_url", image_url = new { url = $"data:{im.Mime};base64,{im.Base64}" } });
            }
            payload = new
            {
                model = this._model,
                messages = new object[]
                {
                    new { role = "system", content = systemPrompt },
                    new { role = "user", content = parts }
                }
            };
        }

        string json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { PropertyNamingPolicy = JsonNamingPolicy.CamelCase });
        using StringContent httpContent = new StringContent(json, Encoding.UTF8, "application/json");
        using HttpResponseMessage resp = await this._http.PostAsync("https://api.openai.com/v1/chat/completions", httpContent, ct);
        resp.EnsureSuccessStatusCode();

        using Stream stream = await resp.Content.ReadAsStreamAsync(ct);
        using JsonDocument doc = await JsonDocument.ParseAsync(stream, cancellationToken: ct);
        // Extract content
        JsonElement root = doc.RootElement;
        string raw = root.GetProperty("choices")[0].GetProperty("message").GetProperty("content").GetString() ?? string.Empty;
        raw = StripFences(raw);

        SlideExtractionResult? parsed;
        try
        {
            parsed = JsonSerializer.Deserialize<SlideExtractionResult>(raw, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (Exception ex)
        {
            throw new Exception("Failed to parse LLM JSON output: " + ex.Message + "\nRaw text: " + raw);
        }
        parsed ??= new SlideExtractionResult();

        // Populate usage metrics if present
        if (root.TryGetProperty("usage", out JsonElement usage))
        {
            if (usage.TryGetProperty("prompt_tokens", out JsonElement pt) && pt.TryGetInt32(out int pVal))
            {
                parsed.PromptTokens = pVal;
            }

            if (usage.TryGetProperty("completion_tokens", out JsonElement ctoks) && ctoks.TryGetInt32(out int cVal))
            {
                parsed.CompletionTokens = cVal;
            }

            if (usage.TryGetProperty("total_tokens", out JsonElement tt) && tt.TryGetInt32(out int tVal))
            {
                parsed.TotalTokens = tVal;
            }
            else if (parsed.PromptTokens > 0 && parsed.CompletionTokens > 0)
            {
                parsed.TotalTokens = parsed.PromptTokens + parsed.CompletionTokens;
            }
        }

        return parsed;
    }

    private static string StripFences(string text)
    {
        string t = text.Trim();
        if (!t.StartsWith("```"))
        {
            return t;
        }

        int first = t.IndexOf('\n');
        int last = t.LastIndexOf("```");
        if (first >= 0 && last > first)
        {
            return t[(first + 1)..last].Trim();
        }

        return t;
    }
}
