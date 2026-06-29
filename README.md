# NoobNotFound.Sheets

NoobNotFound.Sheets is a .NET library that turns a Google Sheet into a lightweight database. A generic
`DataBaseManager<T>` (behind an `IDataBaseManager<T>` interface) gives you CRUD operations against any
model type, with column mapping via attributes, retry handling, local offline caching, and a design meant
to be run as a long-lived singleton in real applications. *The majority of this project was developed with
the assistance of AI.*

---

## Features

- **Generic Data Model Support** — perform CRUD operations on any model type.
- **Column Mapping via Attributes** — map model properties to Google Sheets columns with `[SheetColumn]`, or exclude them with `[SheetIgnore]`.
- **Compiled Property Accessors** — column-to-property mapping is computed once per model type using compiled expression-tree delegates, not `PropertyInfo.GetValue/SetValue` on every cell.
- **Header Support** — automatically recognizes or adds a header row.
- **Singleton-Safe by Design** — ships an `IDataBaseManager<T>` abstraction and a constructor that accepts a shared `SheetsService`, so multiple models against the same spreadsheet don't each spin up their own `HttpClient`.
- **Conflict-Free Execution** — a `SemaphoreSlim` serializes writes against a given sheet (see [Singleton Usage](#singleton-usage--dependency-injection) — this only works correctly if the manager is registered as a singleton).
- **Local CSV Caching** — work offline and sync later. Cache updates patch the existing snapshot instead of re-fetching the whole sheet on every write.
- **Memory Caching** — faster reads by keeping frequently-used data in memory.
- **Bulk Operations** — add multiple records in a single batch call.
- **Targeted Updates** — `UpdateAsync` rewrites only the rows that actually changed, not the entire sheet.
- **Retry Policies** — automatic retries (Polly) for genuinely transient errors (HTTP 429/5xx, network failures), with a configurable delay and max-attempt count that are actually honored.
- **Cancellation Support** — every async method accepts an optional `CancellationToken`.
- **Pagination** — fetch data in pages for large datasets.
- **Asynchronous, Lazily-Initialized** — no blocking network calls in the constructor; safe to construct as part of DI container startup.
- **Multi-Model Support** — manage multiple data models in a single spreadsheet.
- **Interactive Testing** — a console sample app for exploring every feature.

---

## Installation

<details>
  <summary>Currently NuGet is unsupported</summary>

You can install the NoobNotFound.Sheets package via NuGet Package Manager:

```
Install-Package NoobNotFound.Sheets
```

Or via the .NET CLI:

```
dotnet add package NoobNotFound.Sheets
```
</details>

1. **Clone the Repository**: Build and use the library directly from the source.
2. **Download DLL**: Obtain precompiled DLLs from the repository's releases and reference them in your project.

---

## Prerequisites

Before using NoobNotFound.Sheets, ensure you have:

1. **Google Cloud Setup**: Create a project and enable the Google Sheets API.
2. **Service Account Credentials**: Download the JSON key file for your service account, and share the target spreadsheet with that service account's email address.

---

## Quick Start

### 1. Define a model

```csharp
public class SampleModel
{
    [SheetColumn(0)]
    public int Id { get; set; }

    [SheetColumn(1)]
    public string Name { get; set; } = string.Empty;

    [SheetColumn(2)]
    public string Description { get; set; } = string.Empty;
}
```

### 2. Create a manager

```csharp
using NoobNotFound.Sheets;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;

var credential = GoogleCredential.FromFile("path/to/credentials.json")
    .CreateScoped(SheetsService.Scope.Spreadsheets);

IDataBaseManager<SampleModel> db = new DataBaseManager<SampleModel>(
    credential,
    spreadsheetId: "your-spreadsheet-id",
    sheetName: "YourSheetName");

// Optional: warm it up explicitly so a missing sheet or bad credential fails now,
// rather than on whatever call happens to run first.
await db.EnsureReadyAsync();
```

This overload owns its own `SheetsService` (and therefore its own `HttpClient`). If you have multiple
models against the same spreadsheet, see [Singleton Usage](#singleton-usage--dependency-injection) for
sharing one `SheetsService` between them instead.

### 3. Use it

```csharp
await db.AddAsync(new SampleModel { Id = 1, Name = "Alice", Description = "First entry" });

var allItems = await db.GetAllAsync();

var matches = await db.SearchAsync(item => item.Name.Contains("Ali"));

await db.UpdateAsync(item => item.Id == 1, new SampleModel { Id = 1, Name = "Alice Smith", Description = "Updated" });

await db.RemoveAsync(item => item.Id == 1);
```

---

## Defining a Data Model

Use `[SheetColumn(index)]` to pin a property to a specific column (0 = column A, 1 = column B, etc.).
Properties without `[SheetColumn]` are filled into whatever columns are left over, in declaration order.
Use `[SheetIgnore]` to exclude a property from the sheet entirely.

```csharp
public class Product
{
    [SheetColumn(0)]
    public string Sku { get; set; } = string.Empty;

    [SheetColumn(1)]
    public decimal Price { get; set; }

    // No [SheetColumn] -- fills the next free column automatically.
    public string Category { get; set; } = string.Empty;

    // Never written to or read from the sheet.
    [SheetIgnore]
    public decimal ComputedDiscount { get; set; }
}
```

Notes:
- Duplicate `[SheetColumn]` indices throw `InvalidOperationException` the first time that model type is touched.
- Mapped properties need both a getter **and** a setter (the column mapping is built from compiled
  expression-tree delegates at type-load time) — mark computed, read-only properties `[SheetIgnore]`.
- Supported property types are written/read natively: `string`, numeric types, `bool`, `enum`, `DateTime`,
  `DateTimeOffset`, `Guid`, and their nullable equivalents. Anything else round-trips through JSON.

---

## CRUD Operations

All async methods accept an optional trailing `CancellationToken ct = default`.

### Adding Data

```csharp
var newItem = new SampleModel { Id = 2, Name = "Bob", Description = "Second entry" };
await db.AddAsync(newItem);
```

### Adding Multiple Items

```csharp
var items = new List<SampleModel>
{
    new() { Id = 3, Name = "Charlie" },
    new() { Id = 4, Name = "Dana" },
};
await db.AddRangeAsync(items);
```

### Retrieving All Data

```csharp
var allItems = await db.GetAllAsync();
```

### Searching Data

```csharp
var results = await db.SearchAsync(item => item.Name.StartsWith("B"));
```

### Updating Data

Only the rows matching the predicate are rewritten on the sheet (a targeted `batchUpdate`), not the whole
sheet:

```csharp
var updatedItem = new SampleModel { Id = 1, Name = "Alice", Description = "Updated" };
await db.UpdateAsync(item => item.Id == 1, updatedItem);
```

### Removing Data

```csharp
await db.RemoveAsync(item => item.Id == 1);
```

### Pagination

```csharp
var (items, totalPages) = await db.GetPageAsync(pageSize: 10, pageNumber: 1);
```

### Cancellation

```csharp
using var cts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
await db.AddAsync(newItem, cts.Token);
```

---

## Local Caching (Offline Support)

```csharp
var options = new DatabaseManagerOptions
{
    EnableLocalCache = true,
    LocalCachePath = "path/to/cache",
    CacheExpiration = TimeSpan.FromMinutes(10),
};

IDataBaseManager<SampleModel> db = new DataBaseManager<SampleModel>(
    credential, spreadsheetId, sheetName, options);
```

When enabled, the full dataset is mirrored to a CSV file at `LocalCachePath` and kept in an in-memory
cache. `GetAllAsync()` prefers the in-memory cache, falls back to the CSV file, and only hits the network
if neither is available. Writes (`AddAsync`, `AddRangeAsync`, `UpdateAsync`, `RemoveAsync`) update this
snapshot directly from data already fetched during the write itself — no extra round trip to Google Sheets
just to keep the cache fresh.

`GetAllAsync(useCache: false)` bypasses the cache entirely and always hits the network.

---

## Singleton Usage & Dependency Injection

**`DataBaseManager<T>` is designed to be a long-lived singleton — one instance per model type, per
spreadsheet+sheet.** This isn't just a performance suggestion: the `SemaphoreSlim` that serializes writes
against a sheet only protects you if every caller is going through the *same* instance. Registering this
as `Scoped` or `Transient` in a DI container gives every resolution its own semaphore, and concurrent
requests can then write to the same sheet at the same time — silently defeating the "conflict-free
execution" guarantee.

To make singleton usage cheap and safe:

- **Construction does no network I/O.** The sheet-metadata fetch and header check are deferred until the
  first real call (or run eagerly via `EnsureReadyAsync()`), so building this as part of DI container
  startup doesn't block anything.
- **`SheetsService` can be shared.** If you have several models against the same spreadsheet, construct
  one `SheetsService` and pass it to each `DataBaseManager<T>` via the dedicated constructor overload,
  instead of letting each manager create its own `HttpClient`.

```csharp
// One shared service for the whole app.
var sheetsService = new SheetsService(new BaseClientService.Initializer
{
    HttpClientInitializer = credential,
    ApplicationName = "MyApp"
});

// Each manager reuses it. This manager will NOT dispose the shared service.
IDataBaseManager<UserModel> userDb = new DataBaseManager<UserModel>(sheetsService, spreadsheetId, "Users");
IDataBaseManager<OrderModel> orderDb = new DataBaseManager<OrderModel>(sheetsService, spreadsheetId, "Orders");
```

### Registering with `Microsoft.Extensions.DependencyInjection`

```csharp
builder.Services.AddSharedSheetsService(credential, "MyApp");

builder.Services.AddSheetsDatabase<UserModel>(spreadsheetId, sheetName: "Users");
builder.Services.AddSheetsDatabase<OrderModel>(spreadsheetId, sheetName: "Orders", configureOptions: o =>
{
    o.EnableLocalCache = true;
    o.LocalCachePath = "./Cache";
});

// Elsewhere:
public class MyService(IDataBaseManager<UserModel> users) { /* ... */ }
```

See `SheetsServiceCollectionExtensions.cs` for the extension methods themselves (kept out of the core
library so consumers who don't use a DI container aren't forced to take a dependency on
`Microsoft.Extensions.DependencyInjection.Abstractions`), and the commented usage block at the bottom of
that file for a full `Program.cs` example, including an `IHostedService` that calls `EnsureReadyAsync()`
on every registered manager at startup so a missing sheet fails at deploy time, not on a user's first
request.

---

## Retry Behavior

Writes and reads are wrapped in a retry policy (via Polly) that only retries:

- Google API errors with a `429 Too Many Requests` or `5xx` status code.
- Network-level failures (`HttpRequestException`).

It deliberately does **not** retry on:

- Programming errors (e.g. a duplicate `[SheetColumn]` index, an invalid argument) — these fail
  immediately rather than retrying a request that can't succeed.
- A caller-requested cancellation — if you pass a `CancellationToken` and cancel it, that cancellation is
  respected rather than retried.

Configure retry behavior via `DatabaseManagerOptions`:

```csharp
var options = new DatabaseManagerOptions
{
    MaxRetries = 5,
    RetryDelay = TimeSpan.FromSeconds(1), // base delay; backs off exponentially from here
};
```

---

## Example Console Application

A console application (`TestConsole`) is included to exercise every feature interactively — add, search,
update, remove, paginate, bulk-add. It also demonstrates:

- Calling `EnsureReadyAsync()` once at startup so initialization failures show up immediately.
- Wiring `Ctrl+C` to a `CancellationTokenSource` so an in-flight call can be cancelled without killing the
  process mid-write.
- Depending on `IDataBaseManager<T>` rather than the concrete class.

---

## Architecture Notes

For anyone extending this library:

- **Column mapping is computed once per model type `T`**, not per row. `BuildColumnMap()` resolves
  `[SheetColumn]`/`[SheetIgnore]` attributes and produces a `Dictionary<int, ColumnAccessor>`, where each
  `ColumnAccessor` holds the `PropertyInfo` plus a compiled `Func<object, object?>` getter and
  `Action<object, object?>` setter (built with `System.Linq.Expressions`). This avoids both repeated
  attribute lookups and the per-call overhead of `PropertyInfo.GetValue`/`SetValue`.
- **Column letters are computed dynamically** (`GetColumnLetter`), so models with more than 26 properties
  work correctly (`AA`, `AB`, ...), rather than being silently capped at column `Z`.
- **`SerializeValue`/`DeserializeValue`** special-case strings, primitives, enums, dates, and GUIDs to
  write/read them as plain values; everything else still round-trips through `System.Text.Json`.
- **`IDataBaseManager<T>`** exists so you can mock it in unit tests and so DI containers resolve against an
  abstraction rather than a concrete type. Note that default parameter values (`ct = default`,
  `useCache = true`) are resolved by the compiler based on the *static type* of the reference you're
  calling through — keep the interface and implementation defaults in sync if you ever change one.

---

## Limitations

- **Data Size**: designed for small to medium-sized datasets. `SearchAsync`/`UpdateAsync`/`RemoveAsync`
  evaluate predicates in memory after pulling matching data, since Google Sheets has no query language to
  push a filter down to — there's no way around fetching what you need to search through.
- **Concurrent Access**: the in-process semaphore prevents two writers *within the same application
  instance* from racing, as long as the manager is registered as a singleton (see
  [Singleton Usage](#singleton-usage--dependency-injection)). It does not coordinate writes across
  multiple separate processes or machines hitting the same sheet.
- **Duplicate Columns**: duplicate `[SheetColumn]` indices on a model throw an exception the first time
  that model type is used.

---

## Contributing

Contributions are welcome! Submit issues, feedback, or pull requests on the GitHub repository.

---

## License

This project is licensed under the GNU License. See the LICENSE file for details.
