using Microsoft.AspNetCore.Mvc;
using OsmoDoc.Pptx;

namespace OsmoDoc.API.Controllers;

[ApiController]
[Route("api/presentations")]
public class PresentationController : ControllerBase
{
    private readonly PptxGenerator _pptxGenerator;

    public PresentationController(PptxGenerator pptxGenerator)
    {
        this._pptxGenerator = pptxGenerator;
    }

    /// <summary>
    /// Accepts free-form text and generates a PPTX presentation.
    /// </summary>
    [HttpPost("generate")]
    public async Task<IActionResult> GeneratePresentation([FromBody] PresentationRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.UserText))
        {
            return this.BadRequest("UserText is required.");
        }

        string pptxPath = await this._pptxGenerator.GeneratePresentationFromTextAsync(request.UserText, ct);

        byte[] bytes = await System.IO.File.ReadAllBytesAsync(pptxPath, ct);
        string fileName = Path.GetFileName(pptxPath);

        return this.File(bytes,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            fileName);
    }

    /// <summary>
    /// Accepts text plus optional images (multipart/form-data) and generates PPTX.
    /// Field names: userText (string), images (IFormFile, multiple allowed)
    /// </summary>
    [HttpPost("generate-with-images")]
    [RequestSizeLimit(25_000_000)]
    public async Task<IActionResult> GenerateWithImages([FromForm] PresentationWithImagesRequest request, CancellationToken ct)
    {
        if (string.IsNullOrWhiteSpace(request.UserText))
        {
            return this.BadRequest("UserText is required.");
        }

        IEnumerable<IFormFile> images = (IEnumerable<IFormFile>)(request.Images ?? new List<IFormFile>());
        string pptxPath = await this._pptxGenerator.GeneratePresentationFromTextAndImagesAsync(request.UserText!, images, ct);

        byte[] bytes = await System.IO.File.ReadAllBytesAsync(pptxPath, ct);
        string fileName = Path.GetFileName(pptxPath);
        return this.File(bytes, "application/vnd.openxmlformats-officedocument.presentationml.presentation", fileName);
    }
}

public class PresentationRequest
{
    public string UserText { get; set; } = "";
}

public class PresentationWithImagesRequest
{
    public string? UserText { get; set; }
    public List<IFormFile>? Images { get; set; }
}
