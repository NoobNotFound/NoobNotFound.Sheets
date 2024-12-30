using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Reflection;
using System.IO;
using CsvHelper;
using System.Globalization;
using Polly;
using Microsoft.Extensions.Caching.Memory;
using System.Text.Json;

namespace NoobNotFound.Sheets;

/// <summary>
/// Configuration options for the DatabaseManager
/// </summary>
public class DatabaseManagerOptions
{
    public bool EnableLocalCache { get; set; } = false;
    public string LocalCachePath { get; set; }
    public TimeSpan CacheExpiration { get; set; } = TimeSpan.FromMinutes(5);
    public int MaxRetries { get; set; } = 3;
    public TimeSpan RetryDelay { get; set; } = TimeSpan.FromSeconds(2);
}

/// <summary>
/// Manages database operations on Google Sheets with support for local caching and advanced features
/// </summary>
public class DataBaseManager<T> : IDisposable where T : class, new()
{
    private readonly SheetsService _service;
    private readonly string _spreadsheetId;
    private readonly string _sheetName;
    private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
    private readonly DatabaseManagerOptions _options;
    private readonly IMemoryCache _cache;
    private readonly IAsyncPolicy _retryPolicy;
    private int _sheetId;
    private bool _disposed;

    /// <summary>
    /// Initializes a new instance of the DataBaseManager class with enhanced features
    /// </summary>
    public DataBaseManager(GoogleCredential credentials, string spreadsheetId, string sheetName, 
        DatabaseManagerOptions options = null, string appName = "NoobNotFound.SheetsService")
    {
        _spreadsheetId = spreadsheetId;
        _sheetName = sheetName;
        _options = options ?? new DatabaseManagerOptions();
        _service = new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credentials,
            ApplicationName = appName
        });

        _cache = new MemoryCache(new MemoryCacheOptions());
        _retryPolicy = Policy
            .Handle<Exception>()
            .WaitAndRetryAsync(_options.MaxRetries, 
                retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)));

        InitializeAsync().GetAwaiter().GetResult();
    }

    private async Task AddHeaderRowAsync()
    {
        var properties = typeof(T).GetProperties()
            .Where(p => p.GetCustomAttribute<SheetColumnAttribute>() != null)
            .OrderBy(p => p.GetCustomAttribute<SheetColumnAttribute>().Index)
         .Concat(typeof(T).GetProperties()
            .Where(p => p.GetCustomAttribute<SheetColumnAttribute>() == null));

        var headerRow = properties.Select(p => p.Name).ToArray();

        var range = $"{_sheetName}!A1:{(char)('A' + headerRow.Count() - 1)}1";
        var valueRange = new ValueRange { Values = new List<IList<object>> { headerRow } };

        var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

        await updateRequest.ExecuteAsync();
    }

    private async Task InitializeAsync()
    {
        var spreadsheet = await _retryPolicy.ExecuteAsync(async () =>
            await _service.Spreadsheets.Get(_spreadsheetId).ExecuteAsync());
        
        var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == _sheetName);
        if (sheet == null)
            throw new Exception($"Sheet '{_sheetName}' not found in the spreadsheet.");

        _sheetId = sheet.Properties.SheetId.Value;

        var range = $"{_sheetName}!A1:Z1";
        var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
        var response = await request.ExecuteAsync();

        if (response.Values == null || response.Values.Count == 0)
        {
            await AddHeaderRowAsync();
        }
        if (_options.EnableLocalCache)
            await InitializeLocalCacheAsync();
    }

    private async Task InitializeLocalCacheAsync()
    {
        if (!Directory.Exists(_options.LocalCachePath))
            Directory.CreateDirectory(_options.LocalCachePath);

        await RefreshLocalCacheAsync();
    }

    private async Task RefreshLocalCacheAsync()
    {
        var data = await GetAllAsync(useCache: false);
        var cacheFilePath = Path.Combine(_options.LocalCachePath, $"{_sheetName}_cache.csv");
        _cache.Set($"{_sheetName}_all_data", data, _options.CacheExpiration);
        using var writer = new StreamWriter(cacheFilePath);
        using var csv = new CsvWriter(writer, CultureInfo.InvariantCulture);
        await csv.WriteRecordsAsync(data);
    }

    /// <summary>
    /// Adds a new item to the sheet with retry logic and cache update
    /// </summary>
    public async Task<bool> AddAsync(T item)
    {
        try
        {
            await _semaphore.WaitAsync();
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var result = await AddInternalAsync(item);
                if (result && _options.EnableLocalCache)
                    await RefreshLocalCacheAsync();
                return result;
            });
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Adds multiple items in a single batch operation
    /// </summary>
    public async Task<bool> AddRangeAsync(IEnumerable<T> items)
    {
        try
        {
            await _semaphore.WaitAsync();
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var values = items.Select(ConvertToRow).ToList();
                var range = $"{_sheetName}!A:Z";
                var valueRange = new ValueRange { Values = values };
                
                var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                
                var response = await request.ExecuteAsync();
                
                if (response.Updates.UpdatedRows > 0 && _options.EnableLocalCache)
                    await RefreshLocalCacheAsync();
                    
                return response.Updates.UpdatedRows > 0;
            });
        }
        finally
        {
            _semaphore.Release();
        }
    }

    /// <summary>
    /// Retrieves a specific page of items
    /// </summary>
    public async Task<(IEnumerable<T> Items, int TotalPages)> GetPageAsync(int pageSize, int pageNumber)
    {
        if (pageSize <= 0 || pageNumber <= 0)
            throw new ArgumentException("Page size and number must be positive.");

        var allData = await GetAllAsync();
        var totalItems = allData.Count();
        var totalPages = (int)Math.Ceiling(totalItems / (double)pageSize);

        return (
            allData.Skip((pageNumber - 1) * pageSize).Take(pageSize),
            totalPages
        );
    }

    /// <summary>
    /// Retrieves all items with optional caching
    /// </summary>
    public async Task<IEnumerable<T>> GetAllAsync(bool useCache = true)
    {
        if (useCache && _options.EnableLocalCache)
        {
            _cache.CreateEntry(_sheetName);
            var cacheKey = $"{_sheetName}_all_data";
            if (_cache.TryGetValue(cacheKey, out IEnumerable<T> cachedData))
                return cachedData;

            var cacheFilePath = Path.Combine(_options.LocalCachePath, $"{_sheetName}_cache.csv");
            if (File.Exists(cacheFilePath))
            {
                using var reader = new StreamReader(cacheFilePath);
                using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
                var data = csv.GetRecords<T>().ToList();
                
                _cache.Set(cacheKey, data, TimeSpan.FromMinutes(5));
                return data;
            }
        }

        return await _retryPolicy.ExecuteAsync(async () =>
        {
            var range = $"{_sheetName}!A:Z";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();
            return response.Values.Skip(1).Select(ConvertToItem);
        });
    }

    /// <summary>
    /// Updates items with retry logic and cache refresh
    /// </summary>
    public async Task<bool> UpdateAsync(Func<T, bool> predicate, T updatedItem)
    {
        try
        {
            await _semaphore.WaitAsync();
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var result = await UpdateInternalAsync(predicate, updatedItem);
                if (result && _options.EnableLocalCache)
                    await RefreshLocalCacheAsync();
                return result;
            });
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
            _semaphore?.Dispose();
            _service?.Dispose();
            (_cache as IDisposable)?.Dispose();
        }

        _disposed = true;
    }

    // Private helper methods remain largely unchanged, just add retry logic where needed
    private async Task<bool> AddInternalAsync(T item) 
    {
        var values = new List<IList<object>> { ConvertToRow(item) };
        var range = $"{_sheetName}!A:Z";
        var valueRange = new ValueRange { Values = values };
        
        var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        
        var response = await request.ExecuteAsync();
        return response.Updates.UpdatedRows > 0;
    }

    private async Task<bool> UpdateInternalAsync(Func<T, bool> predicate, T updatedItem)
    {
        var rows = await GetAllRowsAsync();
        var updatedRows = new List<IList<object>> { rows[0] };
        bool updated = false;

        for (int i = 1; i < rows.Count; i++)
        {
            var item = ConvertToItem(rows[i]);
            if (predicate(item))
            {
                updatedRows.Add(ConvertToRow(updatedItem));
                updated = true;
            }
            else
            {
                updatedRows.Add(rows[i]);
            }
        }

        if (!updated)
            return false;

        var range = $"{_sheetName}!A1:Z{rows.Count}";
        var valueRange = new ValueRange { Values = updatedRows };
        var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
        updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;

        var response = await updateRequest.ExecuteAsync();
        return response.UpdatedRows > 0;
    }

/// <summary>
/// Converts an object to a row representation for the sheet.
/// </summary>
/// <param name="item">The item to be converted.</param>
/// <returns>An IList of objects representing the item's properties.</returns>
    private IList<object> ConvertToRow(T item)
    {
        var row = new List<object>();
        var properties = typeof(T).GetProperties();

        var usedIndices = new HashSet<int>();

        int maxIndex = properties.Count() - 1;

        for (int i = 0; i <= maxIndex; i++)
        {
            row.Add(null);
        }

        foreach (var prop in properties)
        {
            var attribute = (SheetColumnAttribute)prop.GetCustomAttributes(typeof(SheetColumnAttribute), false).FirstOrDefault();
            var ignoreAttribute = (SheetIgnoreAttribute)prop.GetCustomAttributes(typeof(SheetIgnoreAttribute), false).FirstOrDefault();
            
            if (attribute != null && ignoreAttribute == null)
            {
                if (usedIndices.Contains(attribute.Index))
                {
                    throw new InvalidOperationException($"Duplicate SheetColumn attribute value {attribute.Index} detected for property {prop.Name}");
                }

                var value = prop.GetValue(item);
                row[attribute.Index] = JsonSerializer.Serialize(value, prop.PropertyType);
                usedIndices.Add(attribute.Index);
            }
        }
        //support for properties without SheetColumn attribute
        foreach (var prop in properties)
        {
            var attribute = (SheetColumnAttribute)prop.GetCustomAttributes(typeof(SheetColumnAttribute), false).FirstOrDefault();
            var ignoreAttribute = (SheetIgnoreAttribute)prop.GetCustomAttributes(typeof(SheetIgnoreAttribute), false).FirstOrDefault();

            if (attribute == null && ignoreAttribute == null)
            {

                for (int i = 0; i <= maxIndex; i++)
                {
                    if (!usedIndices.Contains(i))
                    {
                        var value = prop.GetValue(item);
                        row[i] = JsonSerializer.Serialize(value, prop.PropertyType);
                        usedIndices.Add(i);
                        break;
                    }
                }

            }
        }

        return row;
    }

   private string LowCaseIfBool(string value)
    {
        if (value == "TRUE" || value == "FALSE")
        return value.ToLower();

        return value;
    }

    /// <summary>
    /// Converts a row from the sheet to an object of type T.
    /// </summary>
    /// <param name="row">The row data from the sheet.</param>
    /// <returns>An object of type T populated with the row data.</returns>
    private T ConvertToItem(IList<object> row)
    {
        var item = new T();
        var properties = typeof(T).GetProperties();

        var usedIndices = new HashSet<int>();

        foreach (var prop in properties)
        {
            var attribute = (SheetColumnAttribute)prop.GetCustomAttributes(typeof(SheetColumnAttribute), false).FirstOrDefault();
            var ignoreAttribute = (SheetIgnoreAttribute)prop.GetCustomAttributes(typeof(SheetIgnoreAttribute), false).FirstOrDefault();

            if (attribute != null && ignoreAttribute == null)
            {
                if (usedIndices.Contains(attribute.Index))
                {
                    throw new InvalidOperationException($"Duplicate SheetColumn attribute value {attribute.Index} detected for property {prop.Name}");
                }

                if (attribute.Index < row.Count)
                {
                    var cellValue = row[attribute.Index];
                    if (cellValue != null)
                    {
                        prop.SetValue(item, JsonSerializer.Deserialize(LowCaseIfBool(cellValue.ToString()), prop.PropertyType));
                    }
                }

                usedIndices.Add(attribute.Index);
            }
        }

        int maxIndex = properties.Count() - 1;

        //support for properties without SheetColumn attribute
        foreach (var prop in properties)
        {
            var attribute = (SheetColumnAttribute)prop.GetCustomAttributes(typeof(SheetColumnAttribute), false).FirstOrDefault();
            var ignoreAttribute = (SheetIgnoreAttribute)prop.GetCustomAttributes(typeof(SheetIgnoreAttribute), false).FirstOrDefault();

            if (attribute == null && ignoreAttribute == null)
            {

                for (int i = 0; i <= maxIndex; i++)
                {
                    if (i < row.Count)
                    {
                        if (!usedIndices.Contains(i))
                        {

                            var cellValue = row[i];
                            if (cellValue != null)
                            {
                                prop.SetValue(item, JsonSerializer.Deserialize(LowCaseIfBool(cellValue.ToString()), prop.PropertyType));
                            }
                            usedIndices.Add(i);
                            break;
                        }
                    }
                }

            }
        }

        return item;
    }


/// <summary>
/// Retrieves all rows from the sheet.
/// </summary>
/// <returns>An IList of IList of objects representing all rows in the sheet.</returns>
    private async Task<IList<IList<object>>> GetAllRowsAsync()
    {
        var range = $"{_sheetName}!A:Z";
        var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
        var response = await request.ExecuteAsync();
        return response.Values;
    }
    private async Task BatchDeleteRowsAsync(List<int> rowIndices)
    {
        var requests = new List<Request>();

        rowIndices.Sort((a, b) => b.CompareTo(a));

        foreach (var index in rowIndices)
        {
            requests.Add(new Request
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
            });
        }

        var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
        {
            Requests = requests
        };

        await _service.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync();
    }
    /// <summary>
    /// Searches for items in the sheet based on a predicate.
    /// </summary>
    /// <param name="predicate">A function to test each item for a condition.</param>
    /// <returns>An IEnumerable of items that satisfy the condition.</returns>
    public async Task<IEnumerable<T>> SearchAsync(Func<T, bool> predicate)
    {
        var allData = await GetAllAsync();
        return allData.Where(predicate);
    }

public async Task<bool> RemoveInternalAsync(Func<T, bool> predicate)
{
await _semaphore.WaitAsync();
var rows = await GetAllRowsAsync();
var indicesToRemove = new List<int>();

            for (int i = 1; i < rows.Count; i++)  // Start from 1 to skip header
            {
                var item = ConvertToItem(rows[i]);
                if (predicate(item))
                {
                    indicesToRemove.Add(i);
                }
            }

            if (!indicesToRemove.Any())
                return false;

            await BatchDeleteRowsAsync(indicesToRemove);

            return true;
        
}
/// <summary>
/// Removes items from the sheet based on a predicate.
/// </summary>
/// <param name="predicate">A function to test each item for a condition.</param>
/// <returns>A boolean indicating whether any items were removed.</returns>
    public async Task<bool> RemoveAsync(Func<T, bool> predicate)
    {
        
        try
        {
            await _semaphore.WaitAsync();
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var result = await RemoveInternalAsync(predicate);
                if (result && _options.EnableLocalCache)
                    await RefreshLocalCacheAsync();
                return result;
            });
        }
        finally
        {
            _semaphore.Release();
        }
        
    }
}
