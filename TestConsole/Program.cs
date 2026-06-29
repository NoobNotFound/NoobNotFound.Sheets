using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using NoobNotFound.Sheets;

// Define a simple model for demonstration
public class SampleModel
{
    [SheetColumn(0)]
    public int Id { get; set; }

    [SheetColumn(1)]
    public string Name { get; set; } = string.Empty;

    [SheetColumn(2)]
    public string Description { get; set; } = string.Empty;
}

// Main program with a user-interactive menu
public class Program
{
    // Depend on the interface, not the concrete type -- makes this swappable for a mock in tests
    // and matches how you'd inject it from a DI container in a real app.
    private static IDataBaseManager<SampleModel> _databaseManager = null!;

    // Tied to Ctrl+C below, so a long-running call (e.g. RemoveAsync on a big sheet) can be cancelled
    // without killing the whole process mid-write.
    private static CancellationTokenSource _cts = new();

    public static async Task Main(string[] args)
    {
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true; // don't let the runtime kill us mid-request; cancel the current operation instead
            Console.WriteLine("\nCancelling current operation...");
            _cts.Cancel();
        };

        Console.WriteLine("Initializing Google Sheets Database Manager...");

        // Configuration and setup
        var credentialsPath = "yourapi.json"; // Replace with your credentials path
        var spreadsheetId = "1y6i38PqQiyUeuzDhk2I6Wh3njfdr-d75qF1JMzq68es"; // Replace with your spreadsheet ID
        var sheetName = "TestDB"; // Replace with your sheet name

        var credentials = GoogleCredential.FromFile(credentialsPath).CreateScoped([SheetsService.Scope.Spreadsheets]);
        var options = new DatabaseManagerOptions
        {
            EnableLocalCache = true,
            LocalCachePath = "./Cache",
            CacheExpiration = TimeSpan.FromMinutes(10),
            MaxRetries = 3,
            RetryDelay = TimeSpan.FromSeconds(2),

            // Optional: return after the local cache is updated, then sync to Google Sheets in the background.
            // EnableQueuedWrites = true,
            // QueuedWriteDelay = TimeSpan.FromSeconds(1),
            // QueuedWriteFailureDelay = TimeSpan.FromSeconds(30)
        };

        // This console only ever talks to one model, so it's fine for the manager to own its own
        // SheetsService. If you had several models against the same spreadsheet, you'd construct one
        // shared SheetsService and pass it to each DataBaseManager<T> instead -- see
        // SheetsServiceCollectionExtensions.cs for the DI-flavoured version of that pattern.
        _databaseManager = new DataBaseManager<SampleModel>(credentials, spreadsheetId, sheetName, options);

        try
        {
            // Fail fast: surface a missing sheet, bad credentials, etc. here -- before the menu even
            // shows up -- instead of on whatever menu option happens to be picked first.
            await _databaseManager.EnsureReadyAsync(_cts.Token);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to initialize: {ex.Message}");
            return;
        }

        try
        {
            await DisplayMenu();
        }
        finally
        {
            _databaseManager.Dispose();
        }
    }

    private static async Task DisplayMenu()
    {
        while (true)
        {
            Console.Clear();
            Console.WriteLine("Google Sheets Database Manager");
            Console.WriteLine("-----------------------------------");
            Console.WriteLine("1. Add an Item");
            Console.WriteLine("2. Add Multiple Items");
            Console.WriteLine("3. View All Items");
            Console.WriteLine("4. Search for an Item");
            Console.WriteLine("5. Update an Item");
            Console.WriteLine("6. Remove an Item");
            Console.WriteLine("7. Paginate Items");
            Console.WriteLine("8. Exit");
            Console.WriteLine("-----------------------------------");
            Console.Write("Choose an option: ");
            var choice = Console.ReadLine();

            try
            {
                switch (choice)
                {
                    case "1":
                        await AddItem();
                        break;
                    case "2":
                        await AddMultipleItems();
                        break;
                    case "3":
                        await ViewAllItems();
                        break;
                    case "4":
                        await SearchItem();
                        break;
                    case "5":
                        await UpdateItem();
                        break;
                    case "6":
                        await RemoveItem();
                        break;
                    case "7":
                        await PaginateItems();
                        break;
                    case "8":
                        Console.WriteLine("Exiting...");
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Press Enter to try again.");
                        Console.ReadLine();
                        break;
                }
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("\nOperation was cancelled.");
                _cts = new CancellationTokenSource(); // get a fresh token for the next operation
                Console.WriteLine("Press Enter to continue...");
                Console.ReadLine();
            }
        }
    }

    private static async Task AddItem()
    {
        Console.WriteLine("\nAdd an Item:");
        Console.Write("Enter ID: ");
        int id = int.Parse(Console.ReadLine() ?? "0");
        Console.Write("Enter Name: ");
        string name = Console.ReadLine() ?? string.Empty;
        Console.Write("Enter Description: ");
        string description = Console.ReadLine() ?? string.Empty;

        var item = new SampleModel { Id = id, Name = name, Description = description };
        var result = await _databaseManager.AddAsync(item, _cts.Token);

        Console.WriteLine(result ? "Item added successfully!" : "Failed to add item.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task AddMultipleItems()
    {
        Console.WriteLine("\nAdd Multiple Items:");
        var items = new List<SampleModel>();

        while (true)
        {
            Console.Write("Enter ID (or press Enter to stop): ");
            var input = Console.ReadLine();
            if (string.IsNullOrEmpty(input)) break;

            int id = int.Parse(input);
            Console.Write("Enter Name: ");
            string name = Console.ReadLine() ?? string.Empty;
            Console.Write("Enter Description: ");
            string description = Console.ReadLine() ?? string.Empty;

            items.Add(new SampleModel { Id = id, Name = name, Description = description });
        }

        var result = await _databaseManager.AddRangeAsync(items, _cts.Token);
        Console.WriteLine(result ? "Items added successfully!" : "Failed to add items.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task ViewAllItems()
    {
        Console.WriteLine("\nAll Items:");
        var items = await _databaseManager.GetAllAsync(ct: _cts.Token);
        foreach (var item in items)
        {
            Console.WriteLine($"ID: {item.Id}, Name: {item.Name}, Description: {item.Description}");
        }
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task SearchItem()
    {
        Console.WriteLine("\nSearch for an Item:");
        Console.Write("Enter search term for Name: ");
        string searchTerm = Console.ReadLine() ?? string.Empty;

        var items = await _databaseManager.SearchAsync(
            i => i.Name.Contains(searchTerm, StringComparison.OrdinalIgnoreCase),
            _cts.Token);

        foreach (var item in items)
        {
            Console.WriteLine($"ID: {item.Id}, Name: {item.Name}, Description: {item.Description}");
        }
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task UpdateItem()
    {
        Console.WriteLine("\nUpdate an Item:");
        Console.Write("Enter the ID of the item to update: ");
        int id = int.Parse(Console.ReadLine() ?? "0");
        Console.Write("Enter new Name: ");
        string newName = Console.ReadLine() ?? string.Empty;
        Console.Write("Enter new Description: ");
        string newDescription = Console.ReadLine() ?? string.Empty;

        var result = await _databaseManager.UpdateAsync(
            i => i.Id == id,
            new SampleModel { Id = id, Name = newName, Description = newDescription },
            _cts.Token);

        Console.WriteLine(result ? "Item updated successfully!" : "Failed to update item.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task RemoveItem()
    {
        Console.WriteLine("\nRemove an Item:");
        Console.Write("Enter the ID of the item to remove: ");
        int id = int.Parse(Console.ReadLine() ?? "0");

        var result = await _databaseManager.RemoveAsync(i => i.Id == id, _cts.Token);
        Console.WriteLine(result ? "Item removed successfully!" : "Failed to remove item.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task PaginateItems()
    {
        Console.WriteLine("\nPaginate Items:");
        Console.Write("Enter page size: ");
        int pageSize = int.Parse(Console.ReadLine() ?? "10");
        Console.Write("Enter page number: ");
        int pageNumber = int.Parse(Console.ReadLine() ?? "1");

        var (items, totalPages) = await _databaseManager.GetPageAsync(pageSize, pageNumber, _cts.Token);
        Console.WriteLine($"Page {pageNumber}/{totalPages}");
        foreach (var item in items)
        {
            Console.WriteLine($"ID: {item.Id}, Name: {item.Name}, Description: {item.Description}");
        }
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }
}
