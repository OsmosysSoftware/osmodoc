using System.Collections.Generic;
using System.Linq;
using System.Text;
using OsmoDoc.Pptx.Models;

namespace OsmoDoc.Pptx.Services;

public class HtmlGeneratorService
{
    public string GenerateHtml(SlideExtractionResult slidesResult) => this.GenerateHtml(slidesResult, null);

    public string GenerateHtml(SlideExtractionResult slidesResult, Dictionary<string, string>? imageMap)
    {
        StringBuilder html = new StringBuilder();
        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html lang=\"en\">");
        html.AppendLine("<head>");
        html.AppendLine("    <meta charset=\"UTF-8\">");
        html.AppendLine("    <title>Generated Presentation</title>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");

        int slideNumber = 1;
        foreach (SlideData slide in slidesResult.Slides)
        {
            // Determine images for this slide
            List<string> validImages = new List<string>();
            if (slide.Images != null && imageMap != null)
            {
                foreach(string imgId in slide.Images.Distinct())
                {
                    if (imageMap.ContainsKey(imgId))
                    {
                        validImages.Add(imgId);
                    }
                }
            }
            
            // If multiple images, split into multiple slides. If 0 or 1, render once.
            int slideIterations = (validImages.Count > 1) ? validImages.Count : 1;

            for (int i = 0; i < slideIterations; i++)
            {
                string layout = slide.Layout?.ToLowerInvariant() ?? "content-slide";

                // Map layout name to data-template-slide number from Template.pptx
                string dataTemplateSlide = layout switch
                {
                    "title-slide" => "1",
                    "content-slide" => "2",
                    "two-content-slide" => "3",
                    "thank-you-slide" => "4",
                    _ => "2" // fallback to content slide
                };

                html.AppendLine($"    <!-- Slide {slideNumber} -->");
                html.AppendLine($"    <div class=\"slide\" data-template-slide=\"{dataTemplateSlide}\">");

                // Title Slide
                if (layout == "title-slide")
                {
                    html.AppendLine($"        <h1>{Escape(slide.Title)}</h1>");
                    if (!string.IsNullOrWhiteSpace(slide.Subtitle))
                    {
                        html.AppendLine($"        <h3>{Escape(slide.Subtitle)}</h3>");
                    }
                }
                // Thank You Slide
                else if (layout == "thank-you-slide")
                {
                    html.AppendLine($"        <h1>{Escape(slide.Title)}</h1>");
                    if (!string.IsNullOrWhiteSpace(slide.Subtitle))
                    {
                        html.AppendLine($"        <h3>{Escape(slide.Subtitle)}</h3>");
                    }
                }
                // Two Content Slide (text + optional images) or Content Slide
                else 
                {
                    html.AppendLine($"        <h1>{Escape(slide.Title)}</h1>");
                    if (!string.IsNullOrWhiteSpace(slide.Subtitle) && layout != "two-content-slide")
                    {
                         // Optional subtitle support for normal content slides if design allows, 
                         // but standard logic usually puts subtitle in <p> or bullets? 
                         // Existing code didn't handle Subtitle for two-content/content except title-slide/thank-you.
                         // Keeping existing behavior mostly, but ensuring images are added.
                    }

                    if (slide.Bullets != null && slide.Bullets.Count > 0)
                    {
                        html.AppendLine("        <ul>");
                        foreach (string bullet in slide.Bullets)
                        {
                            html.AppendLine($"            <li>{Escape(bullet)}</li>");
                        }
                        html.AppendLine("        </ul>");
                    }
                    else if (!string.IsNullOrWhiteSpace(slide.Content))
                    {
                        html.AppendLine($"        <p>{Escape(slide.Content)}</p>");
                    }

                    // Image Handling
                    if (validImages.Count > 0 && imageMap != null)
                    {
                        string imgIdToRender;
                        if (validImages.Count > 1)
                        {
                             // Pick the i-th image
                             imgIdToRender = validImages[i];
                        }
                        else
                        {
                             // Pick the only image
                             imgIdToRender = validImages[0];
                        }
                        
                        if (imageMap.TryGetValue(imgIdToRender, out string? rel))
                        {
                             html.AppendLine($"        <img data-image-id='{Escape(imgIdToRender)}' src='{Escape(rel)}' alt='{Escape(imgIdToRender)}' />");
                        }
                    }
                }

                html.AppendLine("    </div>");
                html.AppendLine();
                slideNumber++;
            }
        }

        html.AppendLine("</body>");
        html.AppendLine("</html>");

        return html.ToString();
    }

    private static string Escape(string? input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return string.Empty;
        }

        return System.Net.WebUtility.HtmlEncode(input);
    }
}
