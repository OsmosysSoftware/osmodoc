using System.Collections.Generic;

namespace OsmoDoc.Pptx.Models;

public class SlideData
{
    public string Layout { get; set; } = "content-slide"; // "title-slide" | "content-slide" | "two-content-slide" | "thank-you-slide"
    public string? Title { get; set; }
    public string? Subtitle { get; set; }
    public string? Content { get; set; }
    public List<string>? Bullets { get; set; }
    public List<string>? Images { get; set; } // image IDs (e.g., img1, img2) referencing persisted assets
}

public class SlideExtractionResult
{
    public List<SlideData> Slides { get; set; } = new();
    // Token usage metrics from the LLM response (0 if not provided)
    public int PromptTokens { get; set; }
    public int CompletionTokens { get; set; }
    public int TotalTokens { get; set; }
}
