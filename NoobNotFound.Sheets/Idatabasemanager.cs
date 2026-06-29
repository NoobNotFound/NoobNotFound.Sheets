using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace NoobNotFound.Sheets;

/// <summary>
/// CRUD operations against a Google Sheet treated as a lightweight database for model type T.
/// Implementations (see <see cref="DataBaseManager{T}"/>) are expected to be used as a singleton --
/// see the implementation's remarks for why.
/// </summary>
public interface IDataBaseManager<T> : IDisposable where T : class, new()
{
    /// <summary>
    /// Ensures the manager is initialized (sheet metadata fetched, header row ensured, local cache
    /// warmed if enabled). All other members call this automatically on first use; call it explicitly
    /// at startup if you want initialization failures to surface before the first real request.
    /// </summary>
    Task EnsureReadyAsync(CancellationToken ct = default);

    /// <summary>Adds a new item to the sheet.</summary>
    Task<bool> AddAsync(T item, CancellationToken ct = default);

    /// <summary>Adds multiple items in a single batch operation.</summary>
    Task<bool> AddRangeAsync(IEnumerable<T> items, CancellationToken ct = default);

    /// <summary>Retrieves a specific page of items.</summary>
    Task<(IEnumerable<T> Items, int TotalPages)> GetPageAsync(int pageSize, int pageNumber, CancellationToken ct = default);

    /// <summary>Retrieves all items, preferring the in-memory/local cache when enabled and warm.</summary>
    Task<IEnumerable<T>> GetAllAsync(bool useCache = true, CancellationToken ct = default);

    /// <summary>Searches for items matching a predicate.</summary>
    Task<IEnumerable<T>> SearchAsync(Func<T, bool> predicate, CancellationToken ct = default);

    /// <summary>Updates items matching a predicate.</summary>
    Task<bool> UpdateAsync(Func<T, bool> predicate, T updatedItem, CancellationToken ct = default);

    /// <summary>Removes items matching a predicate.</summary>
    Task<bool> RemoveAsync(Func<T, bool> predicate, CancellationToken ct = default);
}