using System.ComponentModel.DataAnnotations;

namespace OsmoDoc.API.Models;

public class RevokeTokenRequestDTO
{
    [Required(ErrorMessage = "Token is required")]
    public string Token { get; set; } = string.Empty;
}