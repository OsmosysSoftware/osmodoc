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
            // Two Content Slide (text + optional images)
            else if (layout == "two-content-slide")
            {
                html.AppendLine($"        <h1>{Escape(slide.Title)}</h1>");

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

                if (slide.Images != null && slide.Images.Count > 0 && imageMap != null)
                {
                    foreach (string? imgId in slide.Images.Distinct())
                    {
                        if (imageMap.TryGetValue(imgId, out string? rel))
                        {
                            html.AppendLine($"        <img data-image-id='{Escape(imgId)}' src='{Escape(rel)}' alt='{Escape(imgId)}' />");
                        }
                    }
                }
            }
            // Content Slide
            else // content-slide
            {
                html.AppendLine($"        <h1>{Escape(slide.Title)}</h1>");

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
                if (slide.Images != null && slide.Images.Count > 0 && imageMap != null)
                {
                    foreach (string? imgId in slide.Images.Distinct())
                    {
                        if (imageMap.TryGetValue(imgId, out string? rel))
                        {
                            html.AppendLine($"        <img data-image-id='{Escape(imgId)}' src='{Escape(rel)}' alt='{Escape(imgId)}' />");
                        }
                    }
                }
            }

            html.AppendLine("    </div>");
            html.AppendLine();
            slideNumber++;
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
