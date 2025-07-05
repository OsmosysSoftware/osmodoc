using System.ComponentModel.DataAnnotations;

namespace OsmoDoc.API.Models;

public class LoginRequestDTO
{
    [Required(ErrorMessage = "Email is required")]
    [EmailAddress(ErrorMessage = "Invalid email format")]
    public string Email { get; set; } = string.Empty;
}