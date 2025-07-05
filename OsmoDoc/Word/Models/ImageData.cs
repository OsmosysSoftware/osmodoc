using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace OsmoDoc.Word.Models;

public enum ImageSourceType
{
    Base64 = 0,
    LocalFile = 1,
    Url = 2
}

public class ImageData : IValidatableObject
{
    [Required(ErrorMessage = "Placeholder name is required")]
    public string PlaceholderName { get; set; } = string.Empty;

    [Required(ErrorMessage = "Image source type is required")]
    public ImageSourceType SourceType { get; set; }

    [Required(ErrorMessage = "Image data is required")]
    public string Data { get; set; } = string.Empty; // Can be base64, file path, or URL

    public string? ImageExtension { get; set; } // Required for Base64

    public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
    {
        if (this.SourceType == ImageSourceType.Base64)
        {
            if (string.IsNullOrWhiteSpace(this.ImageExtension))
            {
                yield return new ValidationResult(
                    "Image extension is required for Base64 source type",
                    new[] { nameof(this.ImageExtension) }
                );
            }
        }
    }
}