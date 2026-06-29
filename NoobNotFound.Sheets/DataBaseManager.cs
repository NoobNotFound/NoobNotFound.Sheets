using Google;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using CsvHelper;
using System.Globalization;
using System.Text;
using System.Linq.Expressions;
using Polly;
using Microsoft.Extensions.Caching.Memory;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NoobNotFound.Sheets;

/// <summary>
/// Configuration options for the DatabaseManager
/// </summary>
public class DatabaseManagerOptions
{
    public bool EnableLocalCache { get; set; } = false;
    public string? LocalCachePath { get; set; }
    public TimeSpan CacheExpiration { get; set; } = TimeSpan.FromMinutes(5);
    public int MaxRetries { get; set; } = 3;
    public TimeSpan RetryDelay { get; set; } = TimeSpan.FromSeconds(2);
    public bool EnableQueuedWrites { get; set; } = false;
    public TimeSpan QueuedWriteDelay { get; set; } = TimeSpan.FromSeconds(1);
    public TimeSpan QueuedWriteFailureDelay { get; set; } = TimeSpan.FromSeconds(30);
}

/// <summary>
/// Manages database operations on Google Sheets with support for local caching and advanced features.
///
/// THREAD-SAFETY / LIFETIME: this type is designed to be used as a long-lived singleton (one instance
/// per model type T, per spreadsheet+sheet). Its internal SemaphoreSlim is what serializes concurrent
/// writes against the same sheet -- if you register this as Scoped or Transient in a DI container, every
/// resolution gets its own semaphore and the "conflict-free execution" guarantee silently stops working.
/// </summary>
public class DataBaseManager<T> : IDataBaseManager<T> where T : class, new()
{
    // ---- Per-T reflection metadata, computed once instead of on every row ----
    // Getter/Setter are compiled expression-tree delegates rather than PropertyInfo.GetValue/SetValue --
    // built once per property here, then called like ordinary delegates for every cell of every row.
    private sealed class ColumnAccessor
    {
        public required PropertyInfo Property { get; init; }
        public required Func<object, object?> Getter { get; init; }
        public required Action<object, object?> Setter { get; init; }
    }

    private enum QueuedWriteKind
    {
        AppendRows,
        UpdateRows,
        DeleteRows
    }

    private sealed class QueuedWrite
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public QueuedWriteKind Kind { get; set; }
        public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
        public List<List<string>> Rows { get; set; } = [];
        public List<int> DataRowIndices { get; set; } = [];
    }

    private static readonly PropertyInfo[] _properties = typeof(T).GetProperties();
    private static readonly Dictionary<int, ColumnAccessor> _columnMap = BuildColumnMap();
    private static readonly int _columnCount = _columnMap.Count == 0 ? 1 : Math.Max(_columnMap.Keys.Max() + 1, 1);
    private static readonly string _lastColumnLetter = GetColumnLetter(_columnCount - 1);
    private static readonly JsonSerializerOptions _queueJsonOptions = CreateQueueJsonOptions();

    private readonly SheetsService _service;
    private readonly bool _ownsService;
    private readonly string _spreadsheetId;
    private readonly string _sheetName;
    private readonly string _dataRange;
    private readonly SemaphoreSlim _semaphore = new(1, 1);
    private readonly SemaphoreSlim _initLock = new(1, 1);
    private readonly DatabaseManagerOptions _options;
    private readonly IMemoryCache _cache;
    private readonly IAsyncPolicy _retryPolicy;
    private readonly SemaphoreSlim _queueLock = new(1, 1);
    private readonly SemaphoreSlim _queueSignal = new(0);
    private readonly CancellationTokenSource _queueDrainCts = new();
    private readonly List<QueuedWrite> _queuedWrites = [];
    private Task? _queueDrainTask;
    private int _sheetId;
    private volatile bool _initialized;
    private bool _disposed;

    /// <summary>
    /// Creates a manager that owns its own SheetsService (and therefore its own HttpClient / socket pool).
    /// Fine for a single model. If you have several DataBaseManager&lt;T&gt; instances against the same
    /// spreadsheet, prefer the SheetsService-accepting overload below and share one service between them.
    /// </summary>
    public DataBaseManager(
        GoogleCredential credentials,
        string spreadsheetId,
        string sheetName,
        DatabaseManagerOptions? options = null,
        string appName = "NoobNotFound.SheetsService")
        : this(CreateOwnedService(credentials, appName), spreadsheetId, sheetName, options, ownsService: true)
    {
    }

    /// <summary>
    /// Creates a manager backed by a SheetsService you already own (e.g. registered as a singleton and
    /// shared across every DataBaseManager&lt;T&gt; in your app). This manager will NOT dispose it.
    /// </summary>
    public DataBaseManager(
        SheetsService sheetsService,
        string spreadsheetId,
        string sheetName,
        DatabaseManagerOptions? options = null)
        : this(sheetsService, spreadsheetId, sheetName, options, ownsService: false)
    {
    }

    private DataBaseManager(
        SheetsService sheetsService,
        string spreadsheetId,
        string sheetName,
        DatabaseManagerOptions? options,
        bool ownsService)
    {
        _service = sheetsService;
        _ownsService = ownsService;
        _spreadsheetId = spreadsheetId;
        _sheetName = sheetName;
        _options = options ?? new DatabaseManagerOptions();

        if (_options.EnableLocalCache && string.IsNullOrWhiteSpace(_options.LocalCachePath))
            throw new ArgumentException("LocalCachePath must be set when EnableLocalCache is true.", nameof(options));

        if (_options.EnableQueuedWrites && !_options.EnableLocalCache)
            throw new ArgumentException("EnableLocalCache must be true when EnableQueuedWrites is true.", nameof(options));

        _dataRange = $"{_sheetName}!A:{_lastColumnLetter}";
        _cache = new MemoryCache(new MemoryCacheOptions());

        // Only retry genuinely transient failures. Retrying ArgumentException/InvalidOperationException
        // (e.g. a misconfigured SheetColumn attribute) just adds latency before the same error surfaces.
        // Deliberately NOT retrying TaskCanceledException: that fires on user-requested cancellation too,
        // and we don't want to ignore a caller's CancellationToken.
        _retryPolicy = Policy
            .Handle<GoogleApiException>(IsTransientGoogleApiException)
            .Or<HttpRequestException>()
            .WaitAndRetryAsync(_options.MaxRetries,
                retryAttempt => TimeSpan.FromMilliseconds(
                    _options.RetryDelay.TotalMilliseconds * Math.Pow(2, retryAttempt - 1)));

        // NOTE: deliberately no blocking network I/O here. See EnsureInitializedAsync.
    }

    private static JsonSerializerOptions CreateQueueJsonOptions()
    {
        var options = new JsonSerializerOptions(JsonSerializerDefaults.Web)
        {
            WriteIndented = true
        };
        options.Converters.Add(new JsonStringEnumConverter());
        return options;
    }

    private static SheetsService CreateOwnedService(GoogleCredential credentials, string appName) =>
        new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credentials,
            ApplicationName = appName
        });

    private static bool IsTransientGoogleApiException(GoogleApiException ex) =>
        ex.HttpStatusCode == HttpStatusCode.TooManyRequests || (int)ex.HttpStatusCode >= 500;

    // ------------------------------------------------------------------
    // Lazy, async-safe initialization (replaces the old blocking ctor call)
    // ------------------------------------------------------------------

    /// <summary>
    /// Ensures the manager is initialized (sheet metadata fetched, header row ensured, local cache warmed
    /// if enabled). All public methods call this automatically on first use, so you normally don't need to
    /// call it yourself. Call it explicitly at startup (e.g. from an IHostedService) only if you want
    /// initialization failures -- like a missing sheet -- to surface before the first real request.
    /// </summary>
    public Task EnsureReadyAsync(CancellationToken ct = default) => EnsureInitializedAsync(ct);

    private async Task EnsureInitializedAsync(CancellationToken ct)
    {
        if (_initialized) return;

        await _initLock.WaitAsync(ct);
        try
        {
            if (_initialized) return;
            await InitializeCoreAsync(ct);
            _initialized = true;
        }
        finally
        {
            _initLock.Release();
        }
    }

    private async Task InitializeCoreAsync(CancellationToken ct)
    {
        var spreadsheet = await _retryPolicy.ExecuteAsync(
            async token => await _service.Spreadsheets.Get(_spreadsheetId).ExecuteAsync(token), ct);

        var sheet = spreadsheet.Sheets?.FirstOrDefault(s => s.Properties.Title == _sheetName);
        if (sheet is null)
            throw new InvalidOperationException($"Sheet '{_sheetName}' not found in the spreadsheet.");

        _sheetId = sheet.Properties.SheetId ?? 0;

        var headerRange = $"{_sheetName}!A1:{_lastColumnLetter}1";
        var response = await _retryPolicy.ExecuteAsync(async token =>
            await _service.Spreadsheets.Values.Get(_spreadsheetId, headerRange).ExecuteAsync(token), ct);

        if (response.Values is null || response.Values.Count == 0)
            await AddHeaderRowAsync(ct);

        if (_options.EnableLocalCache)
            await InitializeLocalCacheAsync(ct);
    }

    private async Task AddHeaderRowAsync(CancellationToken ct)
    {
        var headerRow = new object[_columnCount];
        foreach (var kvp in _columnMap)
            headerRow[kvp.Key] = kvp.Value.Property.Name;

        var range = $"{_sheetName}!A1:{_lastColumnLetter}1";
        var valueRange = new ValueRange { Values = new List<IList<object>> { headerRow } };

        var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

        await updateRequest.ExecuteAsync(ct);
    }

    private async Task InitializeLocalCacheAsync(CancellationToken ct)
    {
        if (!Directory.Exists(_options.LocalCachePath))
            Directory.CreateDirectory(_options.LocalCachePath!);

        if (!_options.EnableQueuedWrites)
        {
            var fetched = (await FetchAllFromSheetAsync(ct)).ToList();
            await PersistLocalCacheAsync(fetched);
            return;
        }

        await LoadQueuedWritesAsync(ct);

        List<T> data;
        var cacheFilePath = LocalCacheFilePath;
        if (_queuedWrites.Count > 0 && File.Exists(cacheFilePath))
        {
            data = ReadLocalCacheFile();
        }
        else
        {
            data = (await FetchAllFromSheetAsync(ct)).ToList();
            await ReplayQueuedWritesAsync(data, ct);
        }

        await PersistLocalCacheAsync(data);
        StartQueueDrain();
        SignalQueueDrainIfPending();
    }

    // ------------------------------------------------------------------
    // Public CRUD surface
    // ------------------------------------------------------------------

    /// <summary>
    /// Adds a new item to the sheet. If local caching is enabled, the cached snapshot is updated in place
    /// instead of re-fetching the whole sheet.
    /// </summary>
    public async Task<bool> AddAsync(T item, CancellationToken ct = default)
    {
        await EnsureInitializedAsync(ct);
        await _semaphore.WaitAsync(ct);
        try
        {
            if (_options.EnableQueuedWrites)
            {
                var snapshot = await GetCachedSnapshotAsync(ct);
                snapshot.Add(item);
                await EnqueueQueuedWriteAsync(CreateAppendQueuedWrite([item]), snapshot, ct);
                return true;
            }

            return await _retryPolicy.ExecuteAsync(async token =>
            {
                var added = await AddInternalAsync(item, token);
                if (added && _options.EnableLocalCache)
                {
                    var snapshot = await GetCachedSnapshotAsync(token);
                    snapshot.Add(item);
                    await PersistLocalCacheAsync(snapshot);
                }
                return added;
            }, ct);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Adds multiple items in a single batch operation.
    /// </summary>
    public async Task<bool> AddRangeAsync(IEnumerable<T> items, CancellationToken ct = default)
    {
        var itemList = items as IReadOnlyCollection<T> ?? items.ToList();
        if (itemList.Count == 0) return false;

        await EnsureInitializedAsync(ct);
        await _semaphore.WaitAsync(ct);
        try
        {
            if (_options.EnableQueuedWrites)
            {
                var snapshot = await GetCachedSnapshotAsync(ct);
                snapshot.AddRange(itemList);
                await EnqueueQueuedWriteAsync(CreateAppendQueuedWrite(itemList), snapshot, ct);
                return true;
            }

            return await _retryPolicy.ExecuteAsync(async token =>
            {
                var values = itemList.Select(i => (IList<object>)ConvertToRow(i)).ToList();
                var valueRange = new ValueRange { Values = values };

                var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, _dataRange);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

                var response = await request.ExecuteAsync(token);
                var added = (response.Updates?.UpdatedRows ?? 0) > 0;

                if (added && _options.EnableLocalCache)
                {
                    var snapshot = await GetCachedSnapshotAsync(token);
                    snapshot.AddRange(itemList);
                    await PersistLocalCacheAsync(snapshot);
                }
                return added;
            }, ct);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Retrieves a specific page of items.
    /// </summary>
    public async Task<(IEnumerable<T> Items, int TotalPages)> GetPageAsync(int pageSize, int pageNumber, CancellationToken ct = default)
    {
        if (pageSize <= 0 || pageNumber <= 0)
            throw new ArgumentException("Page size and number must be positive.");

        var allData = (await GetAllAsync(ct: ct)).ToList();
        var totalPages = (int)Math.Ceiling(allData.Count / (double)pageSize);

        return (allData.Skip((pageNumber - 1) * pageSize).Take(pageSize), totalPages);
    }

    /// <summary>
    /// Retrieves all items, preferring the in-memory/local cache when enabled and warm.
    /// </summary>
    public async Task<IEnumerable<T>> GetAllAsync(bool useCache = true, CancellationToken ct = default)
    {
        await EnsureInitializedAsync(ct);

        if (useCache && _options.EnableLocalCache)
        {
            if (_cache.TryGetValue(CacheKey, out List<T>? cached) && cached is not null)
                return cached;

            var cacheFilePath = LocalCacheFilePath;
            if (File.Exists(cacheFilePath))
            {
                using var reader = new StreamReader(cacheFilePath);
                using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
                var data = csv.GetRecords<T>().ToList();
                _cache.Set(CacheKey, data, _options.CacheExpiration);
                return data;
            }

            // Cache enabled but cold (e.g. file was deleted out-of-band): fetch once and reseed.
            var fetched = (await FetchAllFromSheetAsync(ct)).ToList();
            if (_options.EnableQueuedWrites)
                await ReplayQueuedWritesAsync(fetched, ct);

            await PersistLocalCacheAsync(fetched);
            return fetched;
        }

        return await FetchAllFromSheetAsync(ct);
    }

    /// <summary>
    /// Searches for items in the sheet based on a predicate.
    /// </summary>
    public async Task<IEnumerable<T>> SearchAsync(Func<T, bool> predicate, CancellationToken ct = default)
    {
        var allData = await GetAllAsync(ct: ct);
        return allData.Where(predicate);
    }

    /// <summary>
    /// Updates items matching the predicate. Only the matching rows are rewritten on the sheet
    /// (a targeted batchUpdate), rather than rewriting every row.
    /// </summary>
    public async Task<bool> UpdateAsync(Func<T, bool> predicate, T updatedItem, CancellationToken ct = default)
    {
        await EnsureInitializedAsync(ct);
        await _semaphore.WaitAsync(ct);
        try
        {
            if (_options.EnableQueuedWrites)
                return await QueueUpdateAsync(predicate, updatedItem, ct);

            return await _retryPolicy.ExecuteAsync(async token =>
            {
                var (hasChanges, allItems, changes) = await BuildUpdatePlanAsync(predicate, updatedItem, token);
                if (!hasChanges)
                    return false;

                var updated = await ApplyUpdatePlanAsync(changes, token);
                if (updated && _options.EnableLocalCache)
                    await PersistLocalCacheAsync(allItems);

                return updated;
            }, ct);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Removes items from the sheet based on a predicate.
    /// </summary>
    public async Task<bool> RemoveAsync(Func<T, bool> predicate, CancellationToken ct = default)
    {
        await EnsureInitializedAsync(ct);
        await _semaphore.WaitAsync(ct);
        try
        {
            if (_options.EnableQueuedWrites)
                return await QueueRemoveAsync(predicate, ct);

            return await _retryPolicy.ExecuteAsync(async token =>
            {
                var (removed, remaining) = await RemoveInternalAsync(predicate, token);
                if (removed && _options.EnableLocalCache)
                    await PersistLocalCacheAsync(remaining);
                return removed;
            }, ct);
        }
        finally
        {
            _semaphore.Release();
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {
            _queueDrainCts.Cancel();
            try
            {
                _queueDrainTask?.Wait(TimeSpan.FromSeconds(5));
            }
            catch (AggregateException ex) when (ex.InnerExceptions.All(e => e is OperationCanceledException))
            {
            }

            _semaphore.Dispose();
            _initLock.Dispose();
            _queueLock.Dispose();
            _queueSignal.Dispose();
            _queueDrainCts.Dispose();
            if (_ownsService)
                _service.Dispose();
            (_cache as IDisposable)?.Dispose();
        }

        _disposed = true;
    }

    // ------------------------------------------------------------------
    // Internal operations
    // ------------------------------------------------------------------

    private async Task<bool> AddInternalAsync(T item, CancellationToken ct)
    {
        var values = new List<IList<object>> { ConvertToRow(item) };
        var valueRange = new ValueRange { Values = values };

        var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, _dataRange);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        var response = await request.ExecuteAsync(ct);
        return (response.Updates?.UpdatedRows ?? 0) > 0;
    }

    private async Task<bool> AppendRowsAsync(IReadOnlyCollection<List<string>> rows, CancellationToken ct)
    {
        if (rows.Count == 0)
            return false;

        var values = rows.Select(row => (IList<object>)row.Cast<object>().ToList()).ToList();
        var valueRange = new ValueRange { Values = values };

        var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, _dataRange);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

        var response = await request.ExecuteAsync(ct);
        return (response.Updates?.UpdatedRows ?? 0) > 0;
    }

    private async Task<(bool HasChanges, List<T> AllItems, List<ValueRange> Changes)> BuildUpdatePlanAsync(
        Func<T, bool> predicate, T updatedItem, CancellationToken ct)
    {
        var rows = await GetAllRowsAsync(ct);
        var allItems = new List<T>(Math.Max(0, rows.Count - 1));
        var changes = new List<ValueRange>();

        for (int i = 1; i < rows.Count; i++) // skip header
        {
            var item = ConvertToItem(rows[i]);
            if (predicate(item))
            {
                allItems.Add(updatedItem);
                int sheetRow = i + 1; // 1-based row number on the actual sheet
                changes.Add(new ValueRange
                {
                    Range = $"{_sheetName}!A{sheetRow}:{_lastColumnLetter}{sheetRow}",
                    Values = new List<IList<object>> { ConvertToRow(updatedItem) }
                });
            }
            else
            {
                allItems.Add(item);
            }
        }

        return (changes.Count > 0, allItems, changes);
    }

    private async Task<bool> ApplyUpdatePlanAsync(List<ValueRange> changes, CancellationToken ct)
    {
        var batchRequest = new BatchUpdateValuesRequest
        {
            ValueInputOption = "USER_ENTERED",
            Data = changes
        };

        var response = await _service.Spreadsheets.Values.BatchUpdate(batchRequest, _spreadsheetId).ExecuteAsync(ct);
        return (response.TotalUpdatedRows ?? 0) > 0;
    }

    private async Task<(bool Removed, List<T> Remaining)> RemoveInternalAsync(Func<T, bool> predicate, CancellationToken ct)
    {
        var rows = await GetAllRowsAsync(ct);
        var indicesToRemove = new List<int>();
        var remaining = new List<T>(Math.Max(0, rows.Count - 1));

        for (int i = 1; i < rows.Count; i++) // skip header
        {
            var item = ConvertToItem(rows[i]);
            if (predicate(item))
                indicesToRemove.Add(i);
            else
                remaining.Add(item);
        }

        if (indicesToRemove.Count == 0)
            return (false, remaining);

        await BatchDeleteRowsAsync(indicesToRemove, ct);
        return (true, remaining);
    }

    private async Task<bool> QueueUpdateAsync(Func<T, bool> predicate, T updatedItem, CancellationToken ct)
    {
        var snapshot = await GetCachedSnapshotAsync(ct);
        var dataRowIndices = new List<int>();
        var rows = new List<List<string>>();

        for (int i = 0; i < snapshot.Count; i++)
        {
            if (!predicate(snapshot[i]))
                continue;

            snapshot[i] = updatedItem;
            dataRowIndices.Add(i);
            rows.Add(ConvertToQueuedRow(updatedItem));
        }

        if (dataRowIndices.Count == 0)
            return false;

        await EnqueueQueuedWriteAsync(new QueuedWrite
        {
            Kind = QueuedWriteKind.UpdateRows,
            DataRowIndices = dataRowIndices,
            Rows = rows
        }, snapshot, ct);

        return true;
    }

    private async Task<bool> QueueRemoveAsync(Func<T, bool> predicate, CancellationToken ct)
    {
        var snapshot = await GetCachedSnapshotAsync(ct);
        var dataRowIndices = new List<int>();
        var remaining = new List<T>(snapshot.Count);

        for (int i = 0; i < snapshot.Count; i++)
        {
            var item = snapshot[i];
            if (predicate(item))
            {
                dataRowIndices.Add(i);
            }
            else
            {
                remaining.Add(item);
            }
        }

        if (dataRowIndices.Count == 0)
            return false;

        await EnqueueQueuedWriteAsync(new QueuedWrite
        {
            Kind = QueuedWriteKind.DeleteRows,
            DataRowIndices = dataRowIndices
        }, remaining, ct);

        return true;
    }

    private static QueuedWrite CreateAppendQueuedWrite(IEnumerable<T> items) => new()
    {
        Kind = QueuedWriteKind.AppendRows,
        Rows = items.Select(ConvertToQueuedRow).ToList()
    };

    private async Task ApplyQueuedWriteAsync(QueuedWrite write, CancellationToken ct)
    {
        bool applied = write.Kind switch
        {
            QueuedWriteKind.AppendRows => await AppendRowsAsync(write.Rows, ct),
            QueuedWriteKind.UpdateRows => await ApplyQueuedUpdateAsync(write, ct),
            QueuedWriteKind.DeleteRows => await ApplyQueuedDeleteAsync(write, ct),
            _ => throw new InvalidOperationException($"Unsupported queued write kind '{write.Kind}'.")
        };

        if (!applied)
            throw new InvalidOperationException($"Queued write '{write.Id}' did not update any rows.");
    }

    private async Task<bool> ApplyQueuedUpdateAsync(QueuedWrite write, CancellationToken ct)
    {
        if (write.DataRowIndices.Count != write.Rows.Count)
            throw new InvalidOperationException($"Queued update '{write.Id}' has mismatched row data.");

        var changes = new List<ValueRange>(write.Rows.Count);
        for (int i = 0; i < write.Rows.Count; i++)
        {
            int sheetRow = write.DataRowIndices[i] + 2; // header row plus 1-based Sheets row numbers
            changes.Add(new ValueRange
            {
                Range = $"{_sheetName}!A{sheetRow}:{_lastColumnLetter}{sheetRow}",
                Values = new List<IList<object>> { write.Rows[i].Cast<object>().ToList() }
            });
        }

        return await ApplyUpdatePlanAsync(changes, ct);
    }

    private async Task<bool> ApplyQueuedDeleteAsync(QueuedWrite write, CancellationToken ct)
    {
        if (write.DataRowIndices.Count == 0)
            return false;

        await BatchDeleteRowsAsync(write.DataRowIndices.Select(index => index + 1).ToList(), ct);
        return true;
    }

    private async Task BatchDeleteRowsAsync(List<int> rowIndices, CancellationToken ct)
    {
        rowIndices.Sort((a, b) => b.CompareTo(a));

        var requests = rowIndices.Select(index => new Request
        {
            DeleteDimension = new DeleteDimensionRequest
            {
                Range = new DimensionRange
                {
                    SheetId = _sheetId,
                    Dimension = "ROWS",
                    StartIndex = index,
                    EndIndex = index + 1
                }
            }
        }).ToList();

        var batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
        await _service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync(ct);
    }

    private async Task<IList<IList<object>>> GetAllRowsAsync(CancellationToken ct)
    {
        return await _retryPolicy.ExecuteAsync(async token =>
        {
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, _dataRange);
            var response = await request.ExecuteAsync(token);
            return response.Values ?? new List<IList<object>>();
        }, ct);
    }

    private async Task<IEnumerable<T>> FetchAllFromSheetAsync(CancellationToken ct)
    {
        return await _retryPolicy.ExecuteAsync(async token =>
        {
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, _dataRange);
            var response = await request.ExecuteAsync(token);
            return response.Values is null
                ? Enumerable.Empty<T>()
                : response.Values.Skip(1).Select(ConvertToItem);
        }, ct);
    }

    // ------------------------------------------------------------------
    // Local cache helpers (CSV mirror for offline support)
    // ------------------------------------------------------------------

    private string CacheKey => $"{_sheetName}_all_data";

    private string LocalCacheFilePath => Path.Combine(_options.LocalCachePath!, $"{_sheetName}_cache.csv");

    private string QueuedWritesFilePath => Path.Combine(_options.LocalCachePath!, $"{_sheetName}_queue.json");

    /// <summary>
    /// Returns a fresh, mutable copy of the current snapshot (cache, falling back to the CSV file, falling
    /// back to a live fetch). Always returns a NEW list so callers can mutate it without corrupting a list
    /// instance that a concurrent GetAllAsync caller might still be enumerating from the cache.
    /// </summary>
    private async Task<List<T>> GetCachedSnapshotAsync(CancellationToken ct)
    {
        if (_cache.TryGetValue(CacheKey, out List<T>? cached) && cached is not null)
            return new List<T>(cached);

        var cacheFilePath = LocalCacheFilePath;
        if (File.Exists(cacheFilePath))
        {
            return ReadLocalCacheFile();
        }

        var fetched = (await FetchAllFromSheetAsync(ct)).ToList();
        if (_options.EnableQueuedWrites)
            await ReplayQueuedWritesAsync(fetched, ct);

        return fetched;
    }

    private List<T> ReadLocalCacheFile()
    {
        using var reader = new StreamReader(LocalCacheFilePath);
        using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
        return csv.GetRecords<T>().ToList();
    }

    private async Task PersistLocalCacheAsync(List<T> data)
    {
        _cache.Set(CacheKey, data, _options.CacheExpiration);

        if (string.IsNullOrWhiteSpace(_options.LocalCachePath))
            return;

        using var writer = new StreamWriter(LocalCacheFilePath);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
        await csv.WriteRecordsAsync(data);
    }

    // ------------------------------------------------------------------
    // Durable queued write helpers
    // ------------------------------------------------------------------

    private async Task LoadQueuedWritesAsync(CancellationToken ct)
    {
        _queuedWrites.Clear();

        if (!File.Exists(QueuedWritesFilePath))
            return;

        using var stream = new FileStream(QueuedWritesFilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var loaded = await JsonSerializer.DeserializeAsync<List<QueuedWrite>>(stream, _queueJsonOptions, ct);
        if (loaded is not null)
            _queuedWrites.AddRange(loaded);
    }

    private async Task EnqueueQueuedWriteAsync(QueuedWrite write, List<T> updatedSnapshot, CancellationToken ct)
    {
        await _queueLock.WaitAsync(ct);
        try
        {
            _queuedWrites.Add(write);
            await PersistQueuedWritesLockedAsync(ct);
            await PersistLocalCacheAsync(updatedSnapshot);
        }
        finally
        {
            _queueLock.Release();
        }

        SignalQueueDrainIfPending();
    }

    private async Task PersistQueuedWritesLockedAsync(CancellationToken ct)
    {
        var tempPath = QueuedWritesFilePath + ".tmp";
        await using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            await JsonSerializer.SerializeAsync(stream, _queuedWrites, _queueJsonOptions, ct);
        }

        File.Move(tempPath, QueuedWritesFilePath, overwrite: true);
    }

    private void StartQueueDrain()
    {
        if (_queueDrainTask is not null)
            return;

        _queueDrainTask = Task.Run(() => DrainQueuedWritesAsync(_queueDrainCts.Token));
    }

    private void SignalQueueDrainIfPending()
    {
        if (!_options.EnableQueuedWrites || _queuedWrites.Count == 0)
            return;

        _queueSignal.Release();
    }

    private async Task DrainQueuedWritesAsync(CancellationToken ct)
    {
        while (!ct.IsCancellationRequested)
        {
            try
            {
                await _queueSignal.WaitAsync(ct);

                while (!ct.IsCancellationRequested)
                {
                    var write = await PeekQueuedWriteAsync(ct);
                    if (write is null)
                        break;

                    try
                    {
                        await _retryPolicy.ExecuteAsync(async token => await ApplyQueuedWriteAsync(write, token), ct);
                        await RemoveQueuedWriteAsync(write.Id, ct);
                        await Task.Delay(_options.QueuedWriteDelay, ct);
                    }
                    catch (OperationCanceledException) when (ct.IsCancellationRequested)
                    {
                        throw;
                    }
                    catch
                    {
                        await Task.Delay(_options.QueuedWriteFailureDelay, ct);
                    }
                }
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                break;
            }
        }
    }

    private async Task<QueuedWrite?> PeekQueuedWriteAsync(CancellationToken ct)
    {
        await _queueLock.WaitAsync(ct);
        try
        {
            return _queuedWrites.Count == 0 ? null : _queuedWrites[0];
        }
        finally
        {
            _queueLock.Release();
        }
    }

    private async Task RemoveQueuedWriteAsync(string id, CancellationToken ct)
    {
        await _queueLock.WaitAsync(ct);
        try
        {
            var index = _queuedWrites.FindIndex(write => write.Id == id);
            if (index < 0)
                return;

            _queuedWrites.RemoveAt(index);
            await PersistQueuedWritesLockedAsync(ct);
        }
        finally
        {
            _queueLock.Release();
        }
    }

    private void ReplayQueuedWrites(List<T> snapshot)
    {
        foreach (var write in _queuedWrites)
        {
            switch (write.Kind)
            {
                case QueuedWriteKind.AppendRows:
                    snapshot.AddRange(write.Rows.Select(ConvertQueuedRowToItem));
                    break;
                case QueuedWriteKind.UpdateRows:
                    for (int i = 0; i < write.Rows.Count && i < write.DataRowIndices.Count; i++)
                    {
                        int index = write.DataRowIndices[i];
                        if (index >= 0 && index < snapshot.Count)
                            snapshot[index] = ConvertQueuedRowToItem(write.Rows[i]);
                    }
                    break;
                case QueuedWriteKind.DeleteRows:
                    foreach (var index in write.DataRowIndices.OrderByDescending(i => i))
                    {
                        if (index >= 0 && index < snapshot.Count)
                            snapshot.RemoveAt(index);
                    }
                    break;
            }
        }
    }

    private async Task ReplayQueuedWritesAsync(List<T> snapshot, CancellationToken ct)
    {
        await _queueLock.WaitAsync(ct);
        try
        {
            ReplayQueuedWrites(snapshot);
        }
        finally
        {
            _queueLock.Release();
        }
    }

    // ------------------------------------------------------------------
    // Reflection metadata (computed once per T, not per row)
    // ------------------------------------------------------------------

    private static Dictionary<int, ColumnAccessor> BuildColumnMap()
    {
        var map = new Dictionary<int, ColumnAccessor>();
        var attributed = new List<(PropertyInfo Prop, int Index)>();
        var unattributed = new List<PropertyInfo>();

        foreach (var prop in _properties)
        {
            if (prop.GetCustomAttribute<SheetIgnoreAttribute>() is not null)
                continue;

            var columnAttr = prop.GetCustomAttribute<SheetColumnAttribute>();
            if (columnAttr is not null)
                attributed.Add((prop, columnAttr.Index));
            else
                unattributed.Add(prop);
        }

        foreach (var (prop, index) in attributed)
        {
            if (!map.TryAdd(index, CreateAccessor(prop)))
                throw new InvalidOperationException(
                    $"Duplicate SheetColumn attribute value {index} detected for property {prop.Name}");
        }

        int slotCount = Math.Max(_properties.Length, map.Count == 0 ? 0 : map.Keys.Max() + 1);
        int nextFree = 0;
        foreach (var prop in unattributed)
        {
            while (map.ContainsKey(nextFree)) nextFree++;
            if (nextFree >= slotCount) slotCount = nextFree + 1;
            map[nextFree] = CreateAccessor(prop);
            nextFree++;
        }

        return map;
    }

    private static ColumnAccessor CreateAccessor(PropertyInfo prop) => new()
    {
        Property = prop,
        Getter = BuildGetter(prop),
        Setter = BuildSetter(prop)
    };

    /// <summary>
    /// Builds a compiled delegate equivalent to "obj => (object)((T)obj).SomeProperty" -- once compiled,
    /// this runs close to the speed of a direct property access, unlike PropertyInfo.GetValue which goes
    /// through reflection's generic invoke path on every call.
    /// </summary>
    private static Func<object, object?> BuildGetter(PropertyInfo prop)
    {
        var instance = Expression.Parameter(typeof(object), "instance");
        var typedInstance = Expression.Convert(instance, typeof(T));
        var propertyAccess = Expression.Property(typedInstance, prop);
        var boxedResult = Expression.Convert(propertyAccess, typeof(object));
        return Expression.Lambda<Func<object, object?>>(boxedResult, instance).Compile();
    }

    /// <summary>
    /// Builds a compiled delegate equivalent to "(obj, value) => ((T)obj).SomeProperty = (PropType)value".
    /// Throws at type-load time (not per-row) if the property has no setter -- mark such properties with
    /// [SheetIgnore] if you don't want them mapped to a column.
    /// </summary>
    private static Action<object, object?> BuildSetter(PropertyInfo prop)
    {
        var instance = Expression.Parameter(typeof(object), "instance");
        var value = Expression.Parameter(typeof(object), "value");
        var typedInstance = Expression.Convert(instance, typeof(T));
        var propertyAccess = Expression.Property(typedInstance, prop);
        var typedValue = Expression.Convert(value, prop.PropertyType);
        var assign = Expression.Assign(propertyAccess, typedValue);
        return Expression.Lambda<Action<object, object?>>(assign, instance, value).Compile();
    }

    private static string GetColumnLetter(int zeroBasedIndex)
    {
        if (zeroBasedIndex < 0) zeroBasedIndex = 0;
        int n = zeroBasedIndex + 1;
        var sb = new StringBuilder();
        while (n > 0)
        {
            int rem = (n - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.ToString();
    }

    private static object[] ConvertToRow(T item)
    {
        var row = new object[_columnCount];
        for (int i = 0; i < row.Length; i++)
            row[i] = string.Empty;

        foreach (var kvp in _columnMap)
            row[kvp.Key] = SerializeValue(kvp.Value.Getter(item), kvp.Value.Property.PropertyType);

        return row;
    }

    private static List<string> ConvertToQueuedRow(T item) =>
        ConvertToRow(item)
            .Select(value => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty)
            .ToList();

    private static T ConvertQueuedRowToItem(List<string> row) =>
        ConvertToItem(row.Cast<object>().ToList());

    private static T ConvertToItem(IList<object> row)
    {
        var item = new T();

        foreach (var kvp in _columnMap)
        {
            int index = kvp.Key;
            if (index >= row.Count) continue;

            var cellValue = row[index];
            if (cellValue is null) continue;

            var deserialized = DeserializeValue(cellValue.ToString() ?? string.Empty, kvp.Value.Property.PropertyType);
            kvp.Value.Setter(item, deserialized);
        }

        return item;
    }

    /// <summary>
    /// Converts a property value into something safe to put in a Sheets cell. Primitives, strings, enums,
    /// dates, and GUIDs are written as plain/native values. Everything else still round-trips through JSON.
    ///
    /// NOTE: this fixes a real bug in the previous implementation, which ran every value -- including
    /// plain strings -- through JsonSerializer.Serialize. For a string that produces a JSON-quoted string
    /// (e.g. "Alice" becomes the literal text "\"Alice\""), so every string cell ended up wrapped in quote
    /// characters when viewed directly in the spreadsheet. If you have existing data written by the old
    /// version, expect those cells to still contain literal quote characters until rewritten.
    /// </summary>
    private static object SerializeValue(object? value, Type type)
    {
        if (value is null) return string.Empty;

        var underlying = Nullable.GetUnderlyingType(type) ?? type;

        if (underlying == typeof(string)) return value;
        if (underlying.IsEnum) return value.ToString() ?? string.Empty;
        if (underlying == typeof(bool)) return value;
        if (underlying.IsPrimitive || underlying == typeof(decimal)) return value;
        if (underlying == typeof(DateTime)) return ((DateTime)value).ToString("o", CultureInfo.InvariantCulture);
        if (underlying == typeof(DateTimeOffset)) return ((DateTimeOffset)value).ToString("o", CultureInfo.InvariantCulture);
        if (underlying == typeof(Guid)) return value.ToString() ?? string.Empty;

        return JsonSerializer.Serialize(value, type);
    }

    private static object? DeserializeValue(string text, Type type)
    {
        var isNullableValueType = Nullable.GetUnderlyingType(type) is not null;
        var underlying = Nullable.GetUnderlyingType(type) ?? type;

        if (string.IsNullOrEmpty(text))
            return type.IsValueType && !isNullableValueType ? Activator.CreateInstance(type) : null;

        if (underlying == typeof(string)) return text;
        if (underlying.IsEnum) return Enum.Parse(underlying, text, ignoreCase: true);
        if (underlying == typeof(bool)) return bool.Parse(text);
        if (underlying == typeof(Guid)) return Guid.Parse(text);
        if (underlying == typeof(DateTime)) return DateTime.Parse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind);
        if (underlying == typeof(DateTimeOffset)) return DateTimeOffset.Parse(text, CultureInfo.InvariantCulture);
        if (underlying.IsPrimitive || underlying == typeof(decimal))
            return Convert.ChangeType(text, underlying, CultureInfo.InvariantCulture);

        return JsonSerializer.Deserialize(text, type);
    }
}
