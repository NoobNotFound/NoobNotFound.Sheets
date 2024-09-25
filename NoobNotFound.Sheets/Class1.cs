using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace NoobNotFound.Sheets;
    /// <summary>
    /// Manages database operations on Google Sheets for a specific model type.
    /// </summary>
    /// <typeparam name="T">The model type to be managed in the database.</typeparam>
    public class DataBaseManager<T> where T : class, new()
    {
        private readonly SheetsService _service;
        private readonly string _spreadsheetId;
        private readonly string _sheetName;

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
        }

        /// <summary>
        /// Adds a new item to the sheet.
        /// </summary>
        /// <param name="item">The item to be added to the database.</param>
        /// <returns>A boolean indicating whether the operation was successful.</returns>
        public async Task<bool> AddAsync(T item)
        {
            var values = new List<IList<object>> { ConvertToRow(item) };
            var range = $"{_sheetName}!A:Z";

            var valueRange = new ValueRange { Values = values };
            var request = _service.Spreadsheets.Values.Append(valueRange, _spreadsheetId, range);
            request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;

            var response = await request.ExecuteAsync();
            return response.Updates.UpdatedRows > 0;
        }

        /// <summary>
        /// Removes items from the sheet based on a predicate.
        /// </summary>
        /// <param name="predicate">A function to test each item for a condition.</param>
        /// <returns>A boolean indicating whether any items were removed.</returns>
        public async Task<bool> RemoveAsync(Func<T, bool> predicate)
        {
            var allData = await GetAllAsync();
            var itemsToRemove = allData.Where(predicate).ToList();

            if (!itemsToRemove.Any())
                return false;

            var rows = await GetAllRowsAsync();
            var indicesToRemove = new List<int>();

            for (int i = 0; i < rows.Count; i++)
            {
                var item = ConvertToItem(rows[i]);
                if (itemsToRemove.Any(x => x.Equals(item)))
                {
                    indicesToRemove.Add(i + 1); // +1 because sheet rows are 1-indexed
                }
            }

            // Remove rows in reverse order to avoid shifting issues
            for (int i = indicesToRemove.Count - 1; i >= 0; i--)
            {
                await DeleteRowAsync(indicesToRemove[i]);
            }

            return true;
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

            return response.Values.Select(ConvertToItem);
        }

        /// <summary>
        /// Updates existing items in the sheet based on a predicate.
        /// </summary>
        /// <param name="predicate">A function to test each item for a condition.</param>
        /// <param name="updatedItem">The updated item to replace the existing ones.</param>
        /// <returns>A boolean indicating whether any items were updated.</returns>
        public async Task<bool> UpdateAsync(Func<T, bool> predicate, T updatedItem)
        {
            var rows = await GetAllRowsAsync();
            var updatedRows = new List<IList<object>>();
            bool updated = false;

            foreach (var row in rows)
            {
                var item = ConvertToItem(row);
                if (predicate(item))
                {
                    updatedRows.Add(ConvertToRow(updatedItem));
                    updated = true;
                }
                else
                {
                    updatedRows.Add(row);
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
            return typeof(T).GetProperties().Select(p => p.GetValue(item)).Cast<object>().ToList();
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

            for (int i = 0; i < Math.Min(properties.Length, row.Count); i++)
            {
                properties[i].SetValue(item, Convert.ChangeType(row[i], properties[i].PropertyType));
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

        /// <summary>
        /// Deletes a specific row from the sheet.
        /// </summary>
        /// <param name="rowIndex">The index of the row to be deleted.</param>
        private async Task DeleteRowAsync(int rowIndex)
        {
            var requestBody = new Request
            {
                DeleteDimension = new DeleteDimensionRequest
                {
                    Range = new DimensionRange
                    {
                        SheetId = 0, // Assuming it's the first sheet, modify if needed
                        Dimension = "ROWS",
                        StartIndex = rowIndex - 1,
                        EndIndex = rowIndex
                    }
                }
            };

            var deleteRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { requestBody }
            };

            var request = _service.Spreadsheets.BatchUpdate(deleteRequest, _spreadsheetId);
            await request.ExecuteAsync();
        }
    }