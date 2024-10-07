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

namespace NoobNotFound.Sheets
{
    /// <summary>
    /// Manages database operations on Google Sheets for a specific model type.
    /// </summary>
    /// <typeparam name="T">The model type to be managed in the database.</typeparam>
    public class DataBaseManager<T> where T : class, new()
    {
        private readonly SheetsService _service;
        private readonly string _spreadsheetId;
        private readonly string _sheetName;
        private readonly SemaphoreSlim _semaphore = new SemaphoreSlim(1, 1);
        private int _sheetId;

    /// <summary>
    /// Initializes a new instance of the DataBaseManager class.
    /// </summary>
    /// <param name="credentials">The Google API credentials</param>
    /// <param name="spreadsheetId">The ID of the Google Sheets spreadsheet to use as the database.</param>
    /// <param name="sheetName">The name of the specific sheet within the spreadsheet to use.</param>
    /// <param name="appName">The name of the your app (optional).</param>
        public DataBaseManager(GoogleCredential credentials, string spreadsheetId, string sheetName, string appName = "NoobNotFound.SheetsService")
        {
            _spreadsheetId = spreadsheetId;
            _sheetName = sheetName;
            _service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credentials,
                ApplicationName = appName
            });

            InitializeAsync().GetAwaiter().GetResult();
        }

        private async Task InitializeAsync()
        {
            var spreadsheet = await _service.Spreadsheets.Get(_spreadsheetId).ExecuteAsync();
            var sheet = spreadsheet.Sheets.FirstOrDefault(s => s.Properties.Title == _sheetName);

            if (sheet == null)
            {
                throw new Exception($"Sheet '{_sheetName}' not found in the spreadsheet.");
            }

            _sheetId = sheet.Properties.SheetId.Value;

            var range = $"{_sheetName}!A1:Z1";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();

            if (response.Values == null || response.Values.Count == 0)
            {
                await AddHeaderRowAsync();
            }
        }

        private async Task AddHeaderRowAsync()
        {
            var properties = typeof(T).GetProperties()
                .Where(p => p.GetCustomAttribute<SheetColumnAttribute>() != null)
                .OrderBy(p => p.GetCustomAttribute<SheetColumnAttribute>().Index);

            var headerRow = properties.Select(p => p.Name).ToArray();

            var range = $"{_sheetName}!A1:{(char)('A' + headerRow.Count() - 1)}1";
            var valueRange = new ValueRange { Values = new List<IList<object>> { headerRow } };

            var updateRequest = _service.Spreadsheets.Values.Update(valueRange, _spreadsheetId, range);
            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;

            await updateRequest.ExecuteAsync();
        }

        /// <summary>
        /// Adds a new item to the sheet.
        /// </summary>
        /// <param name="item">The item to be added to the database.</param>
        /// <returns>A boolean indicating whether the operation was successful.</returns>
        public async Task<bool> AddAsync(T item)
        {
            try
            {
                await _semaphore.WaitAsync();

                var values = new List<IList<object>> { ConvertToRow(item) };
                var range = $"{_sheetName}!A:Z";

                var valueRange = new ValueRange { Values = values };
                var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
                request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                request.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS;

                var response = await request.ExecuteAsync();
                return response.Updates.UpdatedRows > 0;
            }
            finally
            {
                _semaphore.Release();
            }
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
            finally
            {
                _semaphore.Release();
            }
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

    /// <summary>
    /// Retrieves all items from the sheet.
    /// </summary>
    /// <returns>An IEnumerable of all items in the sheet.</returns>
        public async Task<IEnumerable<T>> GetAllAsync()
        {
            var range = $"{_sheetName}!A:Z";
            var request = _service.Spreadsheets.Values.Get(_spreadsheetId, range);
            var response = await request.ExecuteAsync();

            return response.Values.Skip(1).Select(ConvertToItem);  // Skip header row
        }

    /// <summary>
    /// Updates existing items in the sheet based on a predicate.
    /// </summary>
    /// <param name="predicate">A function to test each item for a condition.</param>
    /// <param name="updatedItem">The updated item to replace the existing ones.</param>
    /// <returns>A boolean indicating whether any items were updated.</returns>
        public async Task<bool> UpdateAsync(Func<T, bool> predicate, T updatedItem)
        {
            try
            {
                await _semaphore.WaitAsync();

                var rows = await GetAllRowsAsync();
                var updatedRows = new List<IList<object>> { rows[0] };  // Keep the header
                bool updated = false;

                for (int i = 1; i < rows.Count; i++)  // Start from 1 to skip header
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
            finally
            {
                _semaphore.Release();
            }
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

            int maxIndex = properties
                .Where(p => p.GetCustomAttributes(typeof(SheetColumnAttribute), false).Length > 0)
                .Max(p => ((SheetColumnAttribute)p.GetCustomAttributes(typeof(SheetColumnAttribute), false).First()).Index);

            for (int i = 0; i <= maxIndex; i++)
            {
                row.Add(null);
            }

            foreach (var prop in properties)
            {
                var attribute = (SheetColumnAttribute)prop.GetCustomAttributes(typeof(SheetColumnAttribute), false).FirstOrDefault();
                if (attribute != null)
                {
                    if (usedIndices.Contains(attribute.Index))
                    {
                        throw new InvalidOperationException($"Duplicate SheetColumn attribute value {attribute.Index} detected for property {prop.Name}");
                    }

                    var value = prop.GetValue(item);
                    row[attribute.Index] = value;
                    usedIndices.Add(attribute.Index);
                }
            }

            return row;
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
                if (attribute != null)
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
                            prop.SetValue(item, Convert.ChangeType(cellValue, prop.PropertyType));
                        }
                    }

                    usedIndices.Add(attribute.Index);
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
    }
}