using System;
using System.Collections.Generic;
using System.Linq;
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
    public string Name { get; set; }

    [SheetColumn(2)]
    public string Description { get; set; }
}

// Main program with a user-interactive menu
public class Program
{
    private static DataBaseManager<SampleModel> _databaseManager;

    public static async Task Main(string[] args)
    {
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
            RetryDelay = TimeSpan.FromSeconds(2)
        };

        _databaseManager = new DataBaseManager<SampleModel>(credentials, spreadsheetId, sheetName, options);

        // Display menu
        await DisplayMenu();
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
    }

    private static async Task AddItem()
    {
        Console.WriteLine("\nAdd an Item:");
        Console.Write("Enter ID: ");
        int id = int.Parse(Console.ReadLine());
        Console.Write("Enter Name: ");
        string name = Console.ReadLine();
        Console.Write("Enter Description: ");
        string description = Console.ReadLine();

        var item = new SampleModel { Id = id, Name = name, Description = description };
        var result = await _databaseManager.AddAsync(item);

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
            string name = Console.ReadLine();
            Console.Write("Enter Description: ");
            string description = Console.ReadLine();

            items.Add(new SampleModel { Id = id, Name = name, Description = description });
        }

        var result = await _databaseManager.AddRangeAsync(items);
        Console.WriteLine(result ? "Items added successfully!" : "Failed to add items.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task ViewAllItems()
    {
        Console.WriteLine("\nAll Items:");
        var items = await _databaseManager.GetAllAsync();
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
        string searchTerm = Console.ReadLine();

        var items = await _databaseManager.SearchAsync(i => i.Name.Contains(searchTerm, StringComparison.OrdinalIgnoreCase));
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
        int id = int.Parse(Console.ReadLine());
        Console.Write("Enter new Name: ");
        string newName = Console.ReadLine();
        Console.Write("Enter new Description: ");
        string newDescription = Console.ReadLine();

        var result = await _databaseManager.UpdateAsync(i => i.Id == id, new SampleModel { Id = id, Name = newName, Description = newDescription });
        Console.WriteLine(result ? "Item updated successfully!" : "Failed to update item.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task RemoveItem()
    {
        Console.WriteLine("\nRemove an Item:");
        Console.Write("Enter the ID of the item to remove: ");
        int id = int.Parse(Console.ReadLine());

        var result = await _databaseManager.RemoveAsync(i => i.Id == id);
        Console.WriteLine(result ? "Item removed successfully!" : "Failed to remove item.");
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }

    private static async Task PaginateItems()
    {
        Console.WriteLine("\nPaginate Items:");
        Console.Write("Enter page size: ");
        int pageSize = int.Parse(Console.ReadLine());
        Console.Write("Enter page number: ");
        int pageNumber = int.Parse(Console.ReadLine());

        var (items, totalPages) = await _databaseManager.GetPageAsync(pageSize, pageNumber);
        Console.WriteLine($"Page {pageNumber}/{totalPages}");
        foreach (var item in items)
        {
            Console.WriteLine($"ID: {item.Id}, Name: {item.Name}, Description: {item.Description}");
        }
        Console.WriteLine("Press Enter to continue...");
        Console.ReadLine();
    }
}
