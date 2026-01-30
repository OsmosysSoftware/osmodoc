using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using HtmlAgilityPack;
using Microsoft.AspNetCore.Hosting;

namespace OsmoDoc.Pptx.Services;

public class PptxService
{
    private readonly IWebHostEnvironment _env;

    public PptxService(IWebHostEnvironment env)
    {
        this._env = env;
    }

    public string GeneratePptxFromHtmlString(string htmlContent, string outputFileName, string? outputFolder = null)
    {
        string pptxTemplate = Path.Combine(this._env.ContentRootPath, "Templates", "Template.pptx");
        string outputDir = outputFolder ?? Path.Combine(this._env.ContentRootPath, "Generated");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, outputFileName);

        // Copy template to output
        File.Copy(pptxTemplate, filePath, true);

        // Parse HTML string
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(htmlContent);
        HtmlNodeCollection slideDivs = htmlDoc.DocumentNode.SelectNodes("//div[@class='slide']")
                        ?? new HtmlNodeCollection(null);

        using (PresentationDocument presDoc = PresentationDocument.Open(filePath, true))
        {
            PresentationPart presPart = presDoc.PresentationPart!;
            if (presPart.Presentation == null)
            {
                throw new InvalidOperationException("Presentation element missing in template.");
            }

            SlideIdList slideIdList = presPart.Presentation.SlideIdList ?? presPart.Presentation.AppendChild(new SlideIdList());
            List<SlideId> templateSlideIds = slideIdList.Elements<SlideId>().ToList();
            List<SlidePart> templateSlideParts = new List<SlidePart>();
            foreach (SlideId? sid in templateSlideIds)
            {
                DocumentFormat.OpenXml.StringValue? rel = sid.RelationshipId;
                if (string.IsNullOrEmpty(rel))
                {
                    continue;
                }

                if (presPart.GetPartById(rel!) is SlidePart sp)
                {
                    templateSlideParts.Add(sp);
                }
            }

            slideIdList.RemoveAllChildren();

            uint slideId = 256U;

            foreach (HtmlNode? slideDiv in slideDivs)
            {
                // Determine which template slide to clone
                int templateIndex = slideDiv.GetAttributeValue("data-template-slide", 1) - 1;
                if (templateIndex < 0 || templateIndex >= templateSlideParts.Count)
                {
                    templateIndex = 0; // fallback to first
                }

                SlidePart templateSlidePart = templateSlideParts[templateIndex];
                SlidePart newSlidePart = CloneSlidePart(presPart, templateSlidePart);

                // Extract text from HTML
                HtmlNode h1 = slideDiv.SelectSingleNode(".//h1");
                HtmlNode h3 = slideDiv.SelectSingleNode(".//h3");
                HtmlNodeCollection listItems = slideDiv.SelectNodes(".//li");
                HtmlNode paragraph = slideDiv.SelectSingleNode(".//p");
                HtmlNodeCollection imgNodes = slideDiv.SelectNodes(".//img") ?? new HtmlNodeCollection(null);

                string title = h1?.InnerText?.Trim() ?? "";
                string subtitle = h3?.InnerText?.Trim() ?? "";
                string content = "";

                if (listItems != null && listItems.Count > 0)
                {
                    content = string.Join("\n", listItems.Select(li => li.InnerText.Trim()));
                }
                else if (paragraph != null)
                {
                    content = paragraph.InnerText.Trim();
                }

                // Fill text placeholders
                FillTextPlaceholders(newSlidePart, title, subtitle, content);

                // Fill image placeholders
                foreach (HtmlNode? img in imgNodes)
                {
                    string src = img.GetAttributeValue("src", "");
                    string placeholderName = img.GetAttributeValue("placeholder", "");
                    if (!string.IsNullOrEmpty(src))
                    {
                        string imagePath = Path.Combine(this._env.ContentRootPath, src.Replace('/', Path.DirectorySeparatorChar));
                        if (File.Exists(imagePath))
                        {
                            FillImagePlaceholder(newSlidePart, imagePath, placeholderName);
                        }
                    }
                }

                slideIdList.Append(new SlideId
                {
                    Id = slideId++,
                    RelationshipId = presPart.GetIdOfPart(newSlidePart)
                });
            }

            presPart.Presentation.Save();
        }

        return filePath;
    }


    public string GeneratePptx(string outputFileName)
    {
        string htmlPath = Path.Combine(this._env.ContentRootPath, "Templates", "template.html");
        string pptxTemplate = Path.Combine(this._env.ContentRootPath, "Templates", "Template.pptx");
        string outputDir = Path.Combine(this._env.ContentRootPath, "Generated");
        Directory.CreateDirectory(outputDir);
        string filePath = Path.Combine(outputDir, outputFileName);

        // Copy template to output
        File.Copy(pptxTemplate, filePath, true);

        // Load HTML
        string htmlContent = File.ReadAllText(htmlPath);
        HtmlDocument htmlDoc = new HtmlDocument();
        htmlDoc.LoadHtml(htmlContent);
        HtmlNodeCollection slideDivs = htmlDoc.DocumentNode.SelectNodes("//div[@class='slide']")
                        ?? new HtmlNodeCollection(null);

        using (PresentationDocument presDoc = PresentationDocument.Open(filePath, true))
        {
            PresentationPart presPart = presDoc.PresentationPart!;
            if (presPart.Presentation == null)
            {
                throw new InvalidOperationException("Presentation element missing in template.");
            }

            SlideIdList slideIdList2 = presPart.Presentation.SlideIdList ?? presPart.Presentation.AppendChild(new SlideIdList());
            List<SlideId> templateSlideIds = slideIdList2.Elements<SlideId>().ToList();
            List<SlidePart> templateSlideParts = new List<SlidePart>();
            foreach (SlideId? sid in templateSlideIds)
            {
                DocumentFormat.OpenXml.StringValue? rel = sid.RelationshipId;
                if (string.IsNullOrEmpty(rel))
                {
                    continue;
                }

                if (presPart.GetPartById(rel!) is SlidePart sp)
                {
                    templateSlideParts.Add(sp);
                }
            }

            slideIdList2.RemoveAllChildren();

            uint slideId = 256U;

            foreach (HtmlNode? slideDiv in slideDivs)
            {
                // Determine which template slide to clone
                int templateIndex = slideDiv.GetAttributeValue("data-template-slide", 1) - 1;
                if (templateIndex < 0 || templateIndex >= templateSlideParts.Count)
                {
                    templateIndex = 0; // fallback
                }

                SlidePart templateSlidePart = templateSlideParts[templateIndex];
                SlidePart newSlidePart = CloneSlidePart(presPart, templateSlidePart);

                // Extract text from HTML
                HtmlNode h1 = slideDiv.SelectSingleNode(".//h1");
                HtmlNode h3 = slideDiv.SelectSingleNode(".//h3");
                HtmlNodeCollection listItems = slideDiv.SelectNodes(".//li");
                HtmlNode paragraph = slideDiv.SelectSingleNode(".//p");
                HtmlNodeCollection imgNodes = slideDiv.SelectNodes(".//img") ?? new HtmlNodeCollection(null);

                string title = h1?.InnerText?.Trim() ?? "";
                string subtitle = h3?.InnerText?.Trim() ?? "";
                string content = "";

                if (listItems != null && listItems.Count > 0)
                {
                    content = string.Join("\n", listItems.Select(li => li.InnerText.Trim()));
                }
                else if (paragraph != null)
                {
                    content = paragraph.InnerText.Trim();
                }

                // Fill text placeholders
                FillTextPlaceholders(newSlidePart, title, subtitle, content);

                // Fill image placeholders
                foreach (HtmlNode? img in imgNodes)
                {
                    string src = img.GetAttributeValue("src", "");
                    string placeholderName = img.GetAttributeValue("placeholder", "");
                    if (!string.IsNullOrEmpty(src))
                    {
                        string imagePath = Path.Combine(this._env.ContentRootPath, src.Replace('/', Path.DirectorySeparatorChar));
                        if (File.Exists(imagePath))
                        {
                            FillImagePlaceholder(newSlidePart, imagePath, placeholderName);
                        }
                    }
                }

                slideIdList2.Append(new SlideId
                {
                    Id = slideId++,
                    RelationshipId = presPart.GetIdOfPart(newSlidePart)
                });
            }

            presPart.Presentation.Save();
        }

        return filePath;
    }

    private static SlidePart CloneSlidePart(PresentationPart presPart, SlidePart templateSlidePart)
    {
        SlidePart newSlidePart = presPart.AddNewPart<SlidePart>();
        using (System.IO.Stream stream = templateSlidePart.GetStream(FileMode.Open))
        {
            newSlidePart.FeedData(stream);
        }

        // copy relationships from template
        foreach (IdPartPair rel in templateSlidePart.Parts)
        {
            if (!newSlidePart.Parts.Any(p => p.RelationshipId == rel.RelationshipId))
            {
                newSlidePart.AddPart(rel.OpenXmlPart, rel.RelationshipId);
            }
        }

        return newSlidePart;
    }

    private static void FillTextPlaceholders(SlidePart slidePart, string title, string subtitle, string content)
    {
        IEnumerable<Shape> shapes = slidePart.Slide?.CommonSlideData?.ShapeTree?.Elements<P.Shape>() ?? Enumerable.Empty<P.Shape>();

        foreach (Shape shape in shapes)
        {
            PlaceholderShape? ph = shape.NonVisualShapeProperties?
                         .ApplicationNonVisualDrawingProperties?
                         .GetFirstChild<P.PlaceholderShape>();

            if (ph == null)
            {
                continue;
            }

            if (ph.Type?.Value == P.PlaceholderValues.Title ||
                ph.Type?.Value == P.PlaceholderValues.CenteredTitle)
            {
                SetShapeText(shape, title);
            }
            else if (ph.Type?.Value == P.PlaceholderValues.SubTitle)
            {
                SetShapeText(shape, subtitle);
            }
            else if (ph.Type?.Value == P.PlaceholderValues.Body ||
                     ph.Index != null)
            {
                SetShapeText(shape, content);
            }
        }
    }

    private static void SetShapeText(P.Shape shape, string htmlContent)
    {
        TextBody? textBody = shape.GetFirstChild<P.TextBody>();
        if (textBody == null)
        {
            return;
        }

        textBody.RemoveAllChildren<A.Paragraph>();

        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(htmlContent);

        // Case 1: Handle <ul><li> lists
        HtmlNodeCollection listItems = doc.DocumentNode.SelectNodes("//li");
        if (listItems != null && listItems.Count > 0)
        {
            foreach (HtmlNode? li in listItems)
            {
                textBody.Append(
                    new A.Paragraph(
                        new A.ParagraphProperties(new A.BulletFont { Typeface = "Arial" }),
                        new A.Run(new A.Text(HtmlEntity.DeEntitize(li.InnerText ?? "")))
                    )
                );
            }
            return;
        }

        // Case 2: Handle <p> with <br/> as bullets
        HtmlNodeCollection paragraphs = doc.DocumentNode.SelectNodes("//p");
        if (paragraphs != null)
        {
            foreach (HtmlNode? p in paragraphs)
            {
                string[] lines = p.InnerHtml.Split("<br/>", StringSplitOptions.RemoveEmptyEntries);
                if (lines.Length > 1)
                {
                    // Treat each <br/> split as a bullet
                    foreach (string line in lines)
                    {
                        textBody.Append(
                            new A.Paragraph(
                                new A.ParagraphProperties(new A.BulletFont { Typeface = "Arial" }),
                                new A.Run(new A.Text(HtmlEntity.DeEntitize(line.Trim())))
                            )
                        );
                    }
                }
                else
                {
                    // Single paragraph, no bullets
                    textBody.Append(
                        new A.Paragraph(
                            new A.Run(new A.Text(HtmlEntity.DeEntitize(p.InnerText.Trim())))
                        )
                    );
                }
            }
            return;
        }

        // Case 3: Fallback plain text (split by \n)
        foreach (string line in htmlContent.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            string decodedLine = HtmlEntity.DeEntitize(line);
            textBody.Append(
                new A.Paragraph(
                    new A.Run(new A.Text(decodedLine))
                )
            );
        }
    }

    /// <summary>
    /// Places an image into a picture placeholder (by placeholder name or index), using CENTER CROP to avoid distortion.
    /// </summary>
    private static void FillImagePlaceholder(SlidePart slidePart, string imagePath, string placeholderName)
    {
        if (slidePart?.Slide?.CommonSlideData?.ShapeTree == null)
        {
            return;
        }

        if (!File.Exists(imagePath))
        {
            return;
        }

        ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (FileStream stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
        {
            imagePart.FeedData(stream);
        }
        string rId = slidePart.GetIdOfPart(imagePart);
        if (string.IsNullOrEmpty(rId))
        {
            return;
        }

        (int W, int H)? imgDim = GetImageDimensions(imagePath);

        System.Collections.Generic.IEnumerable<Shape> shapes = slidePart.Slide.CommonSlideData.ShapeTree.Elements<P.Shape>();

        foreach (Shape shape in shapes)
        {
            PlaceholderShape? ph = shape.NonVisualShapeProperties?
                         .ApplicationNonVisualDrawingProperties?
                         .GetFirstChild<P.PlaceholderShape>();
            NonVisualDrawingProperties? nvdp = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;

            // Match by placeholder type Picture or by Name if provided
            bool isPicturePh = ph?.Type?.Value == P.PlaceholderValues.Picture;

            bool nameMatch = false;
            if (!string.IsNullOrEmpty(placeholderName) && nvdp != null)
            {
                nameMatch = string.Equals(nvdp.Name, placeholderName, StringComparison.OrdinalIgnoreCase);
            }

            if (isPicturePh || nameMatch)
            {
                ShapeProperties? shapeProps = shape.ShapeProperties;
                if (shapeProps == null)
                {
                    continue;
                }

                A.BlipFill blipFill = new A.BlipFill();
                blipFill.Blip = new A.Blip { Embed = rId };

                A.SourceRectangle? srcRect = null;
                if (imgDim != null && shapeProps.Transform2D?.Extents != null)
                {
                    long cx = shapeProps.Transform2D.Extents.Cx ?? 0;
                    long cy = shapeProps.Transform2D.Extents.Cy ?? 0;

                    if (cx > 0 && cy > 0)
                    {
                        double shapeRatio = (double)cx / cy;
                        double imgRatio = (double)imgDim.Value.W / imgDim.Value.H;
                        
                        int cropL = 0, cropT = 0, cropR = 0, cropB = 0;
                        
                        if (imgRatio > shapeRatio) // Image is wider
                        {
                            double targetW = imgDim.Value.H * shapeRatio;
                            double diff = imgDim.Value.W - targetW;
                            double pct = diff / imgDim.Value.W;
                            int cropVal = (int)(pct * 100000 / 2);
                            cropL = cropVal;
                            cropR = cropVal;
                        }
                        else if (imgRatio < shapeRatio) // Image is taller
                        {
                            double targetH = imgDim.Value.W / shapeRatio;
                            double diff = imgDim.Value.H - targetH;
                            double pct = diff / imgDim.Value.H;
                            int cropVal = (int)(pct * 100000 / 2);
                            cropT = cropVal;
                            cropB = cropVal;
                        }

                        if (cropL > 0 || cropR > 0 || cropT > 0 || cropB > 0)
                        {
                            srcRect = new A.SourceRectangle
                            {
                                Left = cropL > 0 ? cropL : null, 
                                Top = cropT > 0 ? cropT : null, 
                                Right = cropR > 0 ? cropR : null, 
                                Bottom = cropB > 0 ? cropB : null
                            };
                        }
                    }
                }

                if (srcRect != null)
                {
                    blipFill.Append(srcRect);
                }

                blipFill.Append(new A.Stretch(new A.FillRectangle()));

                shapeProps.RemoveAllChildren<A.BlipFill>();
                shapeProps.Append(blipFill);
                break;
            }
        }
    }

    private static (int W, int H)? GetImageDimensions(string path)
    {
        try
        {
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read))
            using (BinaryReader br = new BinaryReader(fs))
            {
                byte[] header = br.ReadBytes(8);

                // PNG
                if (header.Length >= 8 && header[0] == 0x89 && header[1] == 0x50 && header[2] == 0x4E && header[3] == 0x47)
                {
                    fs.Seek(8, SeekOrigin.Begin);
                    while (fs.Position < fs.Length)
                    {
                        byte[] lenBytes = br.ReadBytes(4);
                        if (BitConverter.IsLittleEndian) Array.Reverse(lenBytes);
                        uint len = BitConverter.ToUInt32(lenBytes, 0);

                        byte[] typeBytes = br.ReadBytes(4);
                        string type = System.Text.Encoding.ASCII.GetString(typeBytes);

                        if (type == "IHDR")
                        {
                            byte[] wBytes = br.ReadBytes(4);
                            byte[] hBytes = br.ReadBytes(4);
                            if (BitConverter.IsLittleEndian) { Array.Reverse(wBytes); Array.Reverse(hBytes); }
                            return (BitConverter.ToInt32(wBytes, 0), BitConverter.ToInt32(hBytes, 0));
                        }
                        fs.Seek(len + 4, SeekOrigin.Current);
                    }
                }
                // JPEG
                else if (header.Length >= 2 && header[0] == 0xFF && header[1] == 0xD8)
                {
                    fs.Seek(2, SeekOrigin.Begin);
                    while (fs.Position < fs.Length)
                    {
                        byte b = br.ReadByte();
                        if (b != 0xFF) continue;
                        byte marker = br.ReadByte();
                        if (marker == 0x00 || (marker >= 0xD0 && marker <= 0xD9)) continue;

                        byte[] lenBytes = br.ReadBytes(2);
                        if (BitConverter.IsLittleEndian) Array.Reverse(lenBytes);
                        ushort len = BitConverter.ToUInt16(lenBytes, 0);

                        if (marker == 0xC0 || marker == 0xC2)
                        {
                            br.ReadByte();
                            byte[] hBytes = br.ReadBytes(2);
                            byte[] wBytes = br.ReadBytes(2);
                            if (BitConverter.IsLittleEndian) { Array.Reverse(hBytes); Array.Reverse(wBytes); }
                            return (BitConverter.ToUInt16(wBytes, 0), BitConverter.ToUInt16(hBytes, 0));
                        }
                        fs.Seek(len - 2, SeekOrigin.Current);
                    }
                }
            }
        }
        catch { }
        return null;
    }
}
