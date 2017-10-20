# Firefly.SimpleXLS

[![NuGet](https://img.shields.io/nuget/v/Firefly.DependencyInjection.svg)](https://www.nuget.org/packages/Firefly.DependencyInjection)
[![license](https://img.shields.io/github/license/mashape/apistatus.svg)]()

> Simple Excel / .XLSX export and import for .NET Core applications in no time.

```cs
using Firefly.SimpleXls;

public void SaveXls(List<Orders> orders)
{
    Exporter.CreateNew()
        .AddSheet(orders)
        .Export("eshop_orders.xlsx");
}

```
## Features

 * Simplified for ready-to-go data exporting
 * Comprehensive .NET API for exporting data model to XLS
 * Allows multiple sheets in one document
 * Allows creating custom data converters
 * Output to file / stream
 * Value localization
 * Cached type reflection
 * Experimental XLS import feature
 

## Installation

Linux/OSX
```
bash$ dotnet add package Firefly.SimpleXLS
```

Windows
```
PM> Install-Package Firefly.SimpleXLS
```

#### Dependencies

 * Netstandard >= 1.6
 * EPPlus.Core >= 1.5.2


---

## Documentation

### The model

Create a model describing the data you want to export. Each property represents one column in the exported document.

```cs

public class XlsOrderViewModel
{
    public string Code { get; set; }
    public string ArticleName { get; set; }
    public decimal Price { get; set; }
    public DateTime CreateAt { get; set; }
}

```
> Hint: For quick mapping between your original entities and XLS view you can use ie. [the Automapper](https://github.com/AutoMapper/AutoMapper)

#### Data types

Supported primitive types:
 - string
 - char
 - int
 - float
 - decimal
 - long
 - boolean
 - _etc..._

Supported complex types:
 - DateTime
 - TimeSpan
 - Enum
 - Tuple of primitives
 - _Some other objects that can be natively represented by .ToString(), like Guid, Point, etc..._ 


#### Model attributes

```cs

public class XlsOrderViewModel
{
    [XlsHeader(Name = "Eshop order code")]   // Will be dispayed as the header of this column 
    public string Code { get; set; }

    public string ArticleName { get; set; }
    public decimal Price { get; set; }
    public DateTime CreateAt { get; set; }

    [XlsIgnore]                              // This column will not be exported 
    public Guid SomeExternalId { get; set; }
}

```


### The export

There are 3 ways where you can export the data to:

 ```cs

public void SaveXls(List<XlsOrderViewModel> orders)
{
    Exporter.CreateNew()
        .AddSheet(orders)
        .Export("eshop_orders.xlsx");

...

    var file = new FileInfo("path/to/eshop_orders.xlsx");
    Exporter.CreateNew()
        .AddSheet(orders)
        .Export(file);
...

    using (var stream = new MemoryStream())
    {
        Exporter.CreateNew()
        .AddSheet(orders)
        .Export(stream);

        // do something with the stream
    }
}

```


#### Export settings

```cs

public void SaveXls(List<XlsOrderViewModel> orders)
{
    Exporter.CreateNew()
        .AddSheet("Orders from Eshop", orders, 
        settings => {
            settings.OmitEmptyColumns = true, // Colums with no values will be omitted
            settings.UseCulture = new CultureInfo("hu-HU"), // Spefific culture for converters. Will be explained below.
            settings.SheetName = "My customized sheet name" // Human-friendly name of the sheet
            }
        )
        .Export("eshop_orders.xlsx");
}

```


### Custom type mapping

You can add any custom or existing type converter with global scope.


 1. Your custom object:

```cs
public class Driver 
{
    public string Name { get; set; }
    public string Phone { get; set; }
}
```


 2. Create a converter:

```cs
public class DriverValueConverter : IValueConverter
{
        public object Write(object item, Type itemType, CultureInfo culture = null)
        {
            if(typeof(Driver).IsAssignableFrom(itemType)){
                var driver = (Driver)item;
                return driver.Name + " / " + driver.Phone;
            }

            throw new ArgumentException("Cannot parse driver for some reason.");
        }

        public object Read(object item)
        {
            var str = (string)item;
            var parts = str.Split('/');
            if (parts.Length != 2){
                throw new ArgumentException("Cannot parse driver.", nameof(item));
            }

            return new Driver {
                Name = parts[0].Trim(),
                Phone = parts[1].Trim()
            };

            // ToDo better error handling ;)
        }
}
```

3. Register your converter


```cs
public void main()
{
    XlsConverters.UseConverter(typeof(Driver), new DriverValueConverter());
}

```

## The import

> Import is an experimental feature. You can basically import only files with known structure.

```cs

public void LoadXls()
{
    var orders = Importer.Open("eshop_orders.xlsx")
        .ImportAs<XlsOrderViewModel>(
            1,                                  // Index of the sheet based on 1. Optional.
            settings => {
                settings.BreakOnError = true    // Throws exception if some value fails to load,
                settings.HasHeader = true       // If the table has a header to be taken in account
            }
        );
}

```