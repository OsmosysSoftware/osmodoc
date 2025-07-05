using Microsoft.AspNetCore.Mvc;
using Serilog.Events;
using Serilog;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Server.Kestrel.Core;
using Microsoft.OpenApi.Models;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Swashbuckle.AspNetCore.Filters;
using OsmoDoc.Pdf;
using StackExchange.Redis;
using OsmoDoc.API.Models;
using OsmoDoc.Services;
using System.IdentityModel.Tokens.Jwt;

WebApplicationBuilder builder = WebApplication.CreateBuilder(args);

// Controller Services
builder.Services.AddControllers(options => options.Filters.Add(new ProducesAttribute("application/json")))
    .AddJsonOptions(options =>
    {
        options.JsonSerializerOptions.DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull;
        options.JsonSerializerOptions.Converters.Add(new JsonStringEnumConverter());
    });

// Load .env file
string root = Directory.GetCurrentDirectory();
string dotenv = Path.GetFullPath(Path.Combine(root, "..", ".env"));
if (File.Exists(dotenv))
{
    OsmoDoc.API.DotEnv.Load(dotenv);
}
else
{
    throw new FileNotFoundException($".env file not found at path: {dotenv}");
}

// Initialize PDF tool path once at startup
OsmoDocPdfConfig.WkhtmltopdfPath = Path.Combine(
    builder.Environment.WebRootPath,
    builder.Configuration.GetSection("STATIC_FILE_PATHS:HTML_TO_PDF_TOOL").Value!
);

// Register REDIS service
builder.Services.AddSingleton<IConnectionMultiplexer>(
    ConnectionMultiplexer.Connect(Environment.GetEnvironmentVariable("REDIS_URL") ?? throw new Exception("No REDIS URL specified"))
);
builder.Services.AddScoped<IRedisTokenStoreService, RedisTokenStoreService>();

// Configure request size limit (50 MB default)
long requestBodySizeLimitBytes = builder.Configuration.GetValue<long>("CONFIG:REQUEST_BODY_SIZE_LIMIT_BYTES", 50 * 1024 * 1024);

// Configure request size for Kestrel server - ASP.NET Core project templates use Kestrel by default when not hosted with IIS
builder.Services.Configure<KestrelServerOptions>(options =>
{
    options.Limits.MaxRequestBodySize = requestBodySizeLimitBytes;
});

// Configure request size for IIS server
builder.Services.Configure<IISServerOptions>(options =>
{
    options.MaxRequestBodySize = requestBodySizeLimitBytes;
});

// AutoMapper Services
builder.Services.AddAutoMapper(typeof(Program));

// Swagger UI Services
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("v1", new OpenApiInfo { Title = "OsmoDoc API", Version = "v1" });

    options.AddSecurityDefinition("oauth2", new OpenApiSecurityScheme
    {
        Name = "Authorization",
        Description = "Standard Authorization header using the Bearer Scheme (\"bearer {token}\")",
        In = ParameterLocation.Header,
        Type = SecuritySchemeType.ApiKey
    });

    options.OperationFilter<SecurityRequirementsOperationFilter>();
});

// Authentication
builder.Services.AddAuthentication(options =>
{
    options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
})
.AddJwtBearer(options =>
{
    string JWT_KEY = Environment.GetEnvironmentVariable("JWT_KEY") ?? throw new InvalidOperationException("No JWT key specified");

    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuerSigningKey = true,
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(JWT_KEY)),
        ValidateIssuer = false,
        ValidateAudience = false,
        // Following code is to allow us to custom handle expiry
        // Here check expiry as nullable
        ClockSkew = TimeSpan.Zero,
        ValidateLifetime = true,
        LifetimeValidator = (DateTime? notBefore, DateTime? expires, SecurityToken securityToken, TokenValidationParameters validationParameters) =>
        {
            // Clone the validation parameters, and remove the defult lifetime validator
            TokenValidationParameters clonedParameters = validationParameters.Clone();
            clonedParameters.LifetimeValidator = null;

            // If token expiry time is not null, then validate lifetime with skewed clock
            if (expires != null)
            {
                Validators.ValidateLifetime(notBefore, expires, securityToken, clonedParameters);
            }

            return true;
        }
    };

    options.Events = new JwtBearerEvents
    {
        OnTokenValidated = async context =>
        {
            IRedisTokenStoreService tokenStore = context.HttpContext.RequestServices.GetRequiredService<IRedisTokenStoreService>();
            JwtSecurityToken? token = context.SecurityToken as JwtSecurityToken;
            string tokenString = string.Empty;

            if (context.Request.Headers.TryGetValue("Authorization", out Microsoft.Extensions.Primitives.StringValues authHeader) &&
                authHeader.ToString().StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase))
            {
                tokenString = authHeader.ToString().Substring("Bearer ".Length).Trim();
            }

            if (!await tokenStore.IsTokenValidAsync(tokenString, context.HttpContext.RequestAborted))
            {
                context.Fail("Token has been revoked.");
            }
        }
    };
});

// Configure Error Response from Model Validations
builder.Services.AddMvc().ConfigureApiBehaviorOptions(options =>
{
    options.InvalidModelStateResponseFactory = actionContext =>
    {
        return ModelValidationBadRequest.ModelValidationErrorResponse(actionContext);
    };
});

// Logging service Serilogs
builder.Logging.AddSerilog();
Log.Logger = new LoggerConfiguration()
    .WriteTo.File(
    path: "wwwroot/logs/log-.txt",
    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}{NewLine}{NewLine}",
    rollingInterval: RollingInterval.Day,
    restrictedToMinimumLevel: LogEventLevel.Information
    ).CreateLogger();

WebApplication app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment() || app.Environment.IsProduction())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseStaticFiles();

app.UseHttpsRedirection();

app.UseAuthentication();

app.UseAuthorization();

app.MapControllers();

app.Run();
