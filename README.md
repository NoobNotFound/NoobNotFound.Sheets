# NoobNotFound.Sheets

NoobNotFound.Sheets is a .NET library that provides a simple and efficient way to use Google Sheets as a database. It offers a generic `DataBaseManager<T>` class that allows you to perform CRUD (Create, Read, Update, Delete) operations on any model type, making it easy to store and retrieve data using Google Sheets as a backend.

## Features

- Generic implementation for flexibility with any data model
- Column mapping via attributes for more control over Google Sheets columns
- Header support
- Uses `SemaphoreSlim` to avoid conflicts.
- Asynchronous methods for improved performance
- CRUD operations: Add, Remove, Search, Update, and Get All
- Easy setup with Google Sheets API Credidental
- Support for multiple model types in a single Google Sheets document

## Installation

<details>
  <summary>Currently NuGet is unsupported</summary>

  
You can install the NoobNotFound.Sheets package via NuGet Package Manager

```
Install-Package NoobNotFound.Sheets
```

Or via the .NET CLI:

```
dotnet add package NoobNotFound.Sheets
```
</details>

Download the dlls from the resources and add it to your project.

Or you can clone this repository and build your own.

## Prerequisites

Before using NoobNotFound.Sheets, ensure you have:

1. Set up a Google Cloud Project
2. Enabled the Google Sheets API
3. Created and downloaded the credentials (JSON key file) for your service account

## Usage

### Initialization

To use the `DataBaseManager<T>`, first import the namespace:

```csharp
using NoobNotFound.Sheets;
```

Then initialize it with your Google Sheets credentials, spreadsheet ID, and sheet name:

```csharp
using Google.Apis.Auth.OAuth2;

var credential = GoogleCredential.FromFile("path/to/your/credentials.json");
var spreadsheetId = "your-spreadsheet-id";
var sheetName = "YourSheetName";

var dbManager = new DataBaseManager<YourModelType>(credential, spreadsheetId, sheetName);
```

### Adding Data with Column Mapping

You can specify column mappings using the `SheetColumn` attribute to map properties to specific columns in the sheet. This helps you organize your data effectively:

```csharp
public class YourModelType
{
    [SheetColumn(0)] // Maps to column A
    public string Name { get; set; }

    [SheetColumn(1)] // Maps to column B
    public int Age { get; set; }
}
```

To add a new item to the sheet:

```csharp
var newItem = new YourModelType { Name = "John", Age = 30 };
bool success = await dbManager.AddAsync(newItem);
```

### Searching Data

To search for items in the sheet:

```csharp
var results = await dbManager.SearchAsync(item => item.Age > 25);
```

### Updating Data

To update existing items in the sheet:

```csharp
var updatedItem = new YourModelType { Name = "John", Age = 31 };
bool success = await dbManager.UpdateAsync(item => item.Name == "John", updatedItem);
```

### Removing Data

To remove items from the sheet:

```csharp
bool success = await dbManager.RemoveAsync(item => item.Name == "John");
```

### Getting All Data

To retrieve all items from the sheet:

```csharp
var allItems = await dbManager.GetAllAsync();
```

### Conflict Handling

If multiple properties in your model are mapped to the same column using the `SheetColumn` attribute, an `InvalidOperationException` will be thrown. For example:

```csharp
public class MyModel
{
    [SheetColumn(0)]
    public string Name { get; set; }

    [SheetColumn(0)]  // This will cause a conflict
    public int Age { get; set; }
}
```

This will result in the following exception:

```plaintext
InvalidOperationException: Duplicate SheetColumn attribute value 0 detected for property Age
```

Ensure each property in your model has a unique column mapping.

## Example

Here's a complete example of how to use NoobNotFound.Sheets with a `Product` model:

```csharp
using NoobNotFound.Sheets;

public class Product
{
    [SheetColumn(0)]
    public string Id { get; set; }

    [SheetColumn(1)]
    public string Name { get; set; }

    [SheetColumn(2)]
    public decimal Price { get; set; }
}

// Initialize the database manager
var productManager = new DataBaseManager<Product>(credential, "spreadsheet-id", "Products");

// Add a new product
var newProduct = new Product { Id = "P001", Name = "Widget", Price = 19.99m };
await productManager.AddAsync(newProduct);

// Search for products
var affordableProducts = await productManager.SearchAsync(p => p.Price < 50);

// Update a product
var updatedProduct = new Product { Id = "P001", Name = "Super Widget", Price = 24.99m };
await productManager.UpdateAsync(p => p.Id == "P001", updatedProduct);

// Remove a product
await productManager.RemoveAsync(p => p.Id == "P001");

// Get all products
var allProducts = await productManager.GetAllAsync();
```

## Limitations

- This library is designed for relatively small to medium-sized datasets. For large amounts of data, consider using a more robust database solution.
- The current implementation doesn't handle concurrent access to the Google Sheet. Use caution in multi-user scenarios.
- Duplicate column mappings will raise exceptions to ensure data integrity.

## Contributing

Contributions to improve NoobNotFound.Sheets are welcome. Please feel free to submit issues or pull requests on our GitHub repository.

## License

This project is licensed under the GNU License - see the LICENSE file for details.
