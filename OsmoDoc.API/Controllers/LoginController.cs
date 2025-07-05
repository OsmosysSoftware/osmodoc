using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using OsmoDoc.API.Models;
using OsmoDoc.API.Helpers;
using OsmoDoc.Services;

namespace OsmoDoc.API.Controllers;

[Route("api")]
[ApiController]
public class LoginController : ControllerBase
{
    private readonly IRedisTokenStoreService _tokenStoreService;
    private readonly ILogger<LoginController> _logger;

    public LoginController(IRedisTokenStoreService tokenStoreService, ILogger<LoginController> logger)
    {
        this._tokenStoreService = tokenStoreService;
        this._logger = logger;
    }

    [HttpPost]
    [Route("login")]
    [AllowAnonymous]
    public async Task<ActionResult<BaseResponse>> Login([FromBody] LoginRequestDTO loginRequest)
    {
        BaseResponse response = new BaseResponse(ResponseStatus.Fail);
        try
        {
            string token = AuthenticationHelper.JwtTokenGenerator(loginRequest.Email);
            await this._tokenStoreService.StoreTokenAsync(token, loginRequest.Email, this.HttpContext.RequestAborted);

            response.Status = ResponseStatus.Success;
            response.AuthToken = token;
            response.Message = "Token generated successfully";
            return this.Ok(response);
        }
        catch (Exception ex)
        {
            response.Status = ResponseStatus.Error;
            response.Message = ex.Message;
            this._logger.LogError(ex.Message);
            this._logger.LogError(ex.StackTrace);
            return this.StatusCode(StatusCodes.Status500InternalServerError, response);
        }
    }

    [HttpPost]
    [Route("revoke")]
    [Authorize]
    public async Task<ActionResult<BaseResponse>> RevokeToken([FromBody] RevokeTokenRequestDTO request)
    {
        BaseResponse response = new BaseResponse(ResponseStatus.Fail);
        try
        {
            await this._tokenStoreService.RevokeTokenAsync(request.Token, this.HttpContext.RequestAborted);

            response.Status = ResponseStatus.Success;
            response.Message = "Token revoked";
            return this.Ok(response);
        }
        catch (Exception ex)
        {
            response.Status = ResponseStatus.Error;
            response.Message = ex.Message;
            this._logger.LogError(ex.Message);
            this._logger.LogError(ex.StackTrace);
            return this.StatusCode(StatusCodes.Status500InternalServerError, response);
        }
    }
}