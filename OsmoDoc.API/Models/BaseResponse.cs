using Microsoft.AspNetCore.Mvc;

namespace OsmoDoc.API.Models;

public enum ResponseStatus
{
    Success,
    Fail,
    Error
}

public class BaseResponse
{
    public BaseResponse(ResponseStatus status) => this.Status = status;
    public ResponseStatus? Status { get; set; }
    public string? Base64 { get; set; }
    public string? AuthToken { get; set; }
    public string? Message { get; set; }
    public string? StackTrace { get; set; }
}

public class ModelValidationBadRequest
{
    public static BadRequestObjectResult ModelValidationErrorResponse(ActionContext actionContext)
    {
        string? firstError = actionContext.ModelState
            .Where(ms => ms.Value != null && ms.Value.Errors.Any())
            .Select(ms => ms.Value!.Errors.FirstOrDefault()?.ErrorMessage)
            .FirstOrDefault(msg => !string.IsNullOrWhiteSpace(msg));

        BaseResponse response = new BaseResponse(ResponseStatus.Error)
        {
            Message = firstError ?? "Validation failed"
        };

        return new BadRequestObjectResult(response);
    }
}
