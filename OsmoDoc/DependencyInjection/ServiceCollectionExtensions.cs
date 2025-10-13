using Microsoft.Extensions.DependencyInjection;
using OsmoDoc.Pptx;
using OsmoDoc.Pptx.Services;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>
/// Helper extensions to wire OsmoDoc services into a dependency injection container.
/// </summary>
public static class OsmoDocServiceCollectionExtensions
{
    /// <summary>
    /// Registers the services required for PPTX generation.
    /// </summary>
    public static IServiceCollection AddOsmoDocPptx(this IServiceCollection services)
    {
        services.AddHttpClient<LlmSlideExtractorService>();
        services.AddSingleton<HtmlGeneratorService>();
        services.AddSingleton<PptxService>();
        services.AddScoped<PptxGenerator>();

        return services;
    }

    /// <summary>
    /// Registers all OsmoDoc services.
    /// Currently includes PPTX generation pipeline.
    /// </summary>
    public static IServiceCollection AddOsmoDocServices(this IServiceCollection services)
        => services.AddOsmoDocPptx();
}
