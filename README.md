# NoobNotFound.Sheets

NoobNotFound.Sheets is a .NET library that simplifies using Google Sheets as a lightweight database. It features a generic `DataBaseManager<T>` class for seamless CRUD (Create, Read, Update, Delete) operations with any model type, along with enhanced caching, bulk operations, and error handling. * ***The majority of this project was developed with the assistance of AI.***

---

## Features

- **Generic Data Model Support**: Perform CRUD operations on any model type.  
- **Column Mapping via Attributes**: Map model properties to Google Sheets columns using `SheetColumn`.  
- **Header Support**: Automatically recognizes or adds headers in the sheet for better organization.  
- **Conflict-Free Execution**: Uses `SemaphoreSlim` to prevent data conflicts during concurrent operations.  
- **Local CSV Caching**: Work offline and sync later using CsvHelper.  
- **Memory Caching**: Faster operations by keeping frequently used data in memory.  
- **Bulk Operations**: Add or update multiple records efficiently in a single operation.  
- **Retry Policies**: Automatic retries with Polly for transient errors.  
- **Pagination**: Fetch data in pages for better performance on large datasets.  
- **Asynchronous Operations**: Non-blocking operations for better performance.  
- **Multi-Model Support**: Manage multiple data models in a single Google Sheets document.  
- **Interactive Testing**: Includes a console sample application for feature exploration.  

---


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

1. **Clone the Repository**: Build and use the library directly from the source.
2. **Download DLL**: Obtain precompiled DLLs from the repository's releases and reference them in your project.

---

## Prerequisites

Before using NoobNotFound.Sheets, ensure you have:

1. **Google Cloud Setup**: Create a project and enable the Google Sheets API.
2. **Service Account Credentials**: Download the JSON key file for your service account.

---

## Usage

### Initialization

Import the namespace and initialize the `DataBaseManager<T>` with your credentials:

```csharp
using NoobNotFound.Sheets;
using Google.Apis.Auth.OAuth2;

var credential = GoogleCredential.FromFile("path/to/credentials.json");
var spreadsheetId = "your-spreadsheet-id";
var sheetName = "YourSheetName";

var dbManager = new DataBaseManager<YourModelType>(credential, spreadsheetId, sheetName);
```

---

### Defining a Data Model

Use the `SheetColumn` attribute to map properties to specific columns:

```csharp
public class YourModelType
{
    [SheetColumn(0)] // Maps to column A
    public string Name { get; set; }

    [SheetColumn(1)] // Maps to column B
    public int Age { get; set; }
}
```

---

### CRUD Operations

#### Adding Data
```csharp
var newItem = new YourModelType { Name = "Alice", Age = 25 };
await dbManager.AddAsync(newItem);
```

#### Searching Data
```csharp
var results = await dbManager.SearchAsync(item => item.Age > 20);
```

#### Updating Data
```csharp
var updatedItem = new YourModelType { Name = "Alice", Age = 26 };
await dbManager.UpdateAsync(item => item.Name == "Alice", updatedItem);
```

#### Removing Data
```csharp
await dbManager.RemoveAsync(item => item.Name == "Alice");
```

#### Retrieving All Data
```csharp
var allItems = await dbManager.GetAllAsync();
```

---

### Advanced Features

#### Bulk Add or Update
```csharp
var items = new List<YourModelType>
{
    new YourModelType { Name = "Bob", Age = 30 },
    new YourModelType { Name = "Charlie", Age = 35 }
};
await dbManager.AddRangeAsync(items);
```

#### Pagination
```csharp
var page = await dbManager.GetPageAsync(pageNumber: 1, pageSize: 10);
```

#### Offline Data with CSV Caching
Enable CSV caching for offline support:
```csharp
var options = new DatabaseManagerOptions
  {
    EnableLocalCache = true,
    LocalCachePath = "path/to/cache",
    CacheExpiration = TimeSpan.FromMinutes(10),
};

dbManager = new DataBaseManager<SampleModel>(credentials, spreadsheetId, sheetName, options);
```

---

## Example Console Application

A console application is included to test all library features interactively. Use the menu-driven interface to perform CRUD operations and explore advanced functionalities like bulk actions and caching.

---

## Limitations

- **Data Size**: Designed for small to medium-sized datasets.
- **Concurrent Access**: Does not support multi-user concurrent operations.
- **Duplicate Columns**: Duplicate mappings in a data model will result in exceptions.

---

## Contributing

Contributions are welcome! Submit issues, feedback, or pull requests on the GitHub repository.

---

## License

This project is licensed under the GNU License. See the LICENSE file for details.
