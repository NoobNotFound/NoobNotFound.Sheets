// Example DI wiring for NoobNotFound.Sheets.
//
// This is intentionally NOT part of the core NoobNotFound.Sheets project -- adding a dependency on
// Microsoft.Extensions.DependencyInjection.Abstractions to the core library would force it on every
// consumer, including plain console apps. Either:
//   (a) drop this file straight into your application project, or
//   (b) put it in a small companion project, e.g. NoobNotFound.Sheets.DependencyInjection, that only
//       depends on Microsoft.Extensions.DependencyInjection.Abstractions (a tiny, dependency-free package).
//
// Requires: NuGet package "Microsoft.Extensions.DependencyInjection.Abstractions"

using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using NoobNotFound.Sheets;

namespace NoobNotFound.Sheets.DependencyInjection;

public static class SheetsServiceCollectionExtensions
{
    /// <summary>
    /// Registers a single shared SheetsService for the app. Call this once. Every model registered via
    /// AddSheetsDatabase below will reuse this same service/HttpClient instead of creating its own.
    /// </summary>
    public static IServiceCollection AddSharedSheetsService(
        this IServiceCollection services,
        GoogleCredential credential,
        string applicationName = "MyApp")
    {
        services.TryAddSingleton(_ => new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = applicationName
        }));

        return services;
    }

    /// <summary>
    /// Registers DataBaseManager&lt;T&gt; as a singleton, backed by the shared SheetsService.
    ///
    /// Must be Singleton, not Scoped/Transient: DataBaseManager&lt;T&gt; holds the semaphore that
    /// serializes writes to this sheet. A new instance per request would mean a new (unshared)
    /// semaphore per request too, and concurrent requests could write to the same sheet at once --
    /// silently defeating the "conflict-free execution" guarantee.
    /// </summary>
    public static IServiceCollection AddSheetsDatabase<T>(
        this IServiceCollection services,
        string spreadsheetId,
        string sheetName,
        Action<DatabaseManagerOptions>? configureOptions = null)
        where T : class, new()
    {
        services.AddSingleton<IDataBaseManager<T>>(sp =>
        {
            var sheetsService = sp.GetRequiredService<SheetsService>();

            var options = new DatabaseManagerOptions();
            configureOptions?.Invoke(options);

            return new DataBaseManager<T>(sheetsService, spreadsheetId, sheetName, options);
        });

        return services;
    }
}

/*
Usage in Program.cs:

var credential = GoogleCredential.FromFile("path/to/credentials.json")
    .CreateScoped(SheetsService.Scope.Spreadsheets);

builder.Services.AddSharedSheetsService(credential, "MyApp");

builder.Services.AddSheetsDatabase<SampleModel>(
    spreadsheetId: "1y6i38PqQiyUeuzDhk2I6Wh3njfdr-d75qF1JMzq68es",
    sheetName: "TestDB",
    configureOptions: o =>
    {
        o.EnableLocalCache = true;
        o.LocalCachePath = "./Cache";
        o.CacheExpiration = TimeSpan.FromMinutes(10);
    });

builder.Services.AddSheetsDatabase<AnotherModel>(spreadsheetId: "...", sheetName: "OtherSheet");

// Elsewhere, inject the interface rather than the concrete type:
//   public class MyService(IDataBaseManager<SampleModel> db) { ... }

// Optional: warm up both managers at startup instead of on first request, so a missing sheet
// or bad credential fails fast during deployment rather than on a user's first call.
public sealed class SheetsWarmupHostedService(
    IDataBaseManager<SampleModel> sampleDb,
    IDataBaseManager<AnotherModel> otherDb) : IHostedService
{
    public Task StartAsync(CancellationToken ct) =>
        Task.WhenAll(sampleDb.EnsureReadyAsync(ct), otherDb.EnsureReadyAsync(ct));

    public Task StopAsync(CancellationToken ct) => Task.CompletedTask;
}

builder.Services.AddHostedService<SheetsWarmupHostedService>();
*/