using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using NoobNotFound.Sheets;

// Sample data model
public class Person
{
    [SheetColumn(0)]
    public string Id { get; set; }

    [SheetColumn(1)]
    public string Name { get; set; }

    [SheetColumn(2)]
    public int Age { get; set; }

    [SheetColumn(3)]
    public string Email { get; set; }

    public override string ToString()
    {
        return $"Id: {Id}, Name: {Name}, Age: {Age}, Email: {Email}";
    }
}

class Program
{
    private static DataBaseManager<Person> _dbManager;

    static async Task Main(string[] args)
    {
        InitializeDatabaseManager();

        while (true)
        {
            DisplayMenu();
            string choice = Console.ReadLine();

            try
            {
                switch (choice)
                {
                    case "1":
                        await AddPerson();
                        break;
                    case "2":
                        await RemovePerson();
                        break;
                    case "3":
                        await SearchPerson();
                        break;
                    case "4":
                        await GetAllPeople();
                        break;
                    case "5":
                        await UpdatePerson();
                        break;
                    case "6":
                        Console.WriteLine("Exiting the application...");
                        return;
                    default:
                        Console.WriteLine("Invalid choice. Please try again.");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            Console.WriteLine("\nPress any key to continue...");
            Console.ReadKey();
            Console.Clear();
        }
    }

    static void InitializeDatabaseManager()
    {
        Console.WriteLine("Initializing Database Manager...");

        // TODO: Replace with your actual spreadsheet ID and sheet name
        string spreadsheetId = "1y6i38PqQiyUeuzDhk2I6Wh3njfdr-d75qF1JMzq68es";
        string sheetName = "TestDB";
        
        _dbManager = new DataBaseManager<Person>(GoogleCredential.FromFile("yourapi.json").CreateScoped([SheetsService.Scope.Spreadsheets]), spreadsheetId, sheetName);
        Console.WriteLine("Database Manager initialized successfully.");
    }

    static void DisplayMenu()
    {
        Console.WriteLine("Google Sheets Database Test Application");
        Console.WriteLine("1. Add a person");
        Console.WriteLine("2. Remove a person");
        Console.WriteLine("3. Search for a person");
        Console.WriteLine("4. Get all people");
        Console.WriteLine("5. Update a person");
        Console.WriteLine("6. Exit");
        Console.Write("Enter your choice (1-6): ");
    }

    static async Task AddPerson()
    {
        Console.WriteLine("\nAdding a new person:");
        var person = new Person
        {
            Id = Guid.NewGuid().ToString(),
            Name = PromptForInput("Enter name: "),
            Age = int.Parse(PromptForInput("Enter age: ")),
            Email = PromptForInput("Enter email: ")
        };

        bool success = await _dbManager.AddAsync(person);
        Console.WriteLine(success ? "Person added successfully." : "Failed to add person.");
    }

    static async Task RemovePerson()
    {
        Console.WriteLine("\nRemoving a person:");
        string idToRemove = PromptForInput("Enter the ID of the person to remove: ");

        bool success = await _dbManager.RemoveAsync(p => p.Id == idToRemove);
        Console.WriteLine(success ? "Person removed successfully." : "Failed to remove person or person not found.");
    }

    static async Task SearchPerson()
    {
        Console.WriteLine("\nSearching for a person:");
        string searchTerm = PromptForInput("Enter a name to search for: ");

        var results = await _dbManager.SearchAsync(p => p.Name.Contains(searchTerm, StringComparison.OrdinalIgnoreCase));

        foreach (var person in results)
        {
            Console.WriteLine(person);
        }
    }

    static async Task GetAllPeople()
    {
        Console.WriteLine("\nGetting all people:");
        var allPeople = await _dbManager.GetAllAsync();

        foreach (var person in allPeople)
        {
            Console.WriteLine(person);
        }
    }

    static async Task UpdatePerson()
    {
        Console.WriteLine("\nUpdating a person:");
        string idToUpdate = PromptForInput("Enter the ID of the person to update: ");

        var personToUpdate = new Person
        {
            Id = idToUpdate,
            Name = PromptForInput("Enter new name: "),
            Age = int.Parse(PromptForInput("Enter new age: ")),
            Email = PromptForInput("Enter new email: ")
        };

        bool success = await _dbManager.UpdateAsync(p => p.Id == idToUpdate, personToUpdate);
        Console.WriteLine(success ? "Person updated successfully." : "Failed to update person or person not found.");
    }

    static string PromptForInput(string prompt)
    {
        Console.Write(prompt);
        return Console.ReadLine();
    }
}