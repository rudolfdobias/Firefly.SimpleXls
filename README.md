# XLS Exporting for .NET Core
> A Firefly.Net package

[![NuGet](https://img.shields.io/nuget/v/Firefly.SimpleXls.svg)](https://www.nuget.org/packages/Firefly.SimpleXls)
[![NuGet](https://img.shields.io/nuget/dt/Firefly.SimpleXls.svg)](https://www.nuget.org/packages/Firefly.SimpleXls)
[![license](https://img.shields.io/github/license/mashape/apistatus.svg)]()

```cs
using Firefly.SimpleXls;

public void ExportToXLS(List<Orders> orders)
{
    Exporter.CreateNew()
        .AddSheet(orders)
        .Export("eshop_orders.xlsx");
}

```
## Features

 * Comprehensive .NET API for simple XLS data exports
 * Allows multiple sheets in one document
 * Allows creating custom data converters
 * Output to file / stream
 * Localization & Translation
 * Cached type reflection
 * Experimental XLS import feature
 * Works on **.NET Core 1.1 or higher**


## Table of contents
<!-- TOC -->

- [XLS Exporting for .NET Core](#xls-exporting-for-net-core)
    - [Features](#features)
    - [Table of contents](#table-of-contents)
    - [Installation](#installation)
        - [Dependencies](#dependencies)
    - [The export](#the-export)
        - [Export settings](#export-settings)
    - [The model](#the-model)
        - [Basic model attributes](#basic-model-attributes)
        - [Data types](#data-types)
    - [Localization & Transalation](#localization--transalation)
        - [Localization](#localization)
            - [Note about data type Localization](#note-about-data-type-localization)
        - [Translation](#translation)
    - [Custom type mapping](#custom-type-mapping)
    - [The import](#the-import)

<!-- /TOC -->

## Installation

Linux/OSX
```
bash$ dotnet add package Firefly.SimpleXLS
```

Windows
```
PM> Install-Package Firefly.SimpleXLS
```

### Dependencies

 * Netstandard >= 1.6
 * EPPlus.Core >= 1.5.2


## The export

 - To a file by string path
 - To a file handle (`FileInfo`)
 - To a `Stream`

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

### Export settings

Use a `SheetExportSettings` action for more detailed configuration.

```cs

public void SaveXls(List<XlsOrderViewModel> orders)
{
    Exporter.CreateNew()
        .AddSheet(orders, 
        settings => {
            settings.OmitEmptyColumns = true,                   // Default true; Colums with no values will be omitted
            settings.SheetName = "My customized sheet name",    // Default model name; Human-friendly name of the sheet
            settings.UseCulture = new CultureInfo("cs-CZ"),     // Default CurrentCulture; Spefific culture for converters and localization.
            settings.Localizer = MyStringLocalizer,             // Default null; Provide an ILocalizer if you want to translate sheet data
            settings.Translate = true                           // Default true; Translates headers with Localizer, if present
            }
        )
        .Export("eshop_orders.xlsx");
}

```

## The model

Create a view model describing the data you want to export. Each property represents one column in the exported document.

```cs

public class XlsOrderViewModel
{
    public string Code { get; set; }
    public string ArticleName { get; set; }
    public decimal Price { get; set; }
    public DateTime CreateAt { get; set; }
}

```
> Hint: For quick mapping between your original entities and XLS view you can use ie. the [Automapper](https://github.com/AutoMapper/AutoMapper)


### Basic model attributes

```cs

[XlsSheet(Name = "My exported orders")]     // Custom sheet name
public class XlsOrderViewModel
{
    [XlsHeader(Name = "Eshop order code")]  // Custom header name
    public string Code { get; set; }

    public string ArticleName { get; set; }
    public decimal Price { get; set; }
    public string CategoryName { get; set; }
    public DateTime CreateAt { get; set; }

    [XlsIgnore]                              // This column will not be exported 
    public Guid SomeExternalId { get; set; }
}

```

### Data types

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

## Localization & Transalation

Sheet name, column headers and custom values can be localized and translated.

### Localization

 - `CultureInfo.CurrentCulture` is taken into account by default.
 - For using custom culture, provide `settings.UseCulture` when initiating export.

#### Note about data type Localization

> Only `DateTime` and `TimeSpan` values are localized to the specified Culture. If you want to auto-localize other types, you may implement own `IValueConverter`.
Localizing other types like _int : 1000.123 => 1,000.123_ is not recommended since Excel handles these datatypes by its own.


```cs
public void CreateExport(List<Order> orders, IStringLocalizer<MyDictionary> myLocalizer)
{
    Exporter.CreateNew()
        .AddSheet(orders, 
        settings => {
            settings.UseCulture = new CultureInfo("cs-CZ")     // Default CurrentCulture; Spefific culture for converters and localization.
            }
        )
        .Export("eshop_orders.xlsx");
}
```

### Translation

 - The sheet is not translated by default until you provide an `IStringLocalizer` when initiating import in `settings.Localizer`.
 - If `settings.Localizer` is provided, all fields including the sheet name are translated by key which is equal as field name.
 - Table cells are **not** translated even if Localizer is present. Use `XlsTranslate` attribute on the column instead.
 - Automatic translation can be turned off setting `settings.Translate` to `FALSE`.

```cs
public void CreateExport(List<Order> orders, IStringLocalizer<MyDictionary> myLocalizer)
{
    Exporter.CreateNew()
        .AddSheet(orders, 
        settings => {
            settings.Localizer = myLocalizer,                  // Default null; Provide an ILocalizer if you want to translate sheet data
            settings.UseCulture = new CultureInfo("cs-CZ")     // Default CurrentCulture;
            }
        )
        .Export("eshop_orders.xlsx");
}
```
```cs

[XlsSheet(Name = "OrderSheetName", DictionaryPrefix="my.dictionary.section.")]
public class XlsOrderViewModel
{
    public string Code { get; set; }

    public string ArticleName { get; set; }
    public decimal Price { get; set; }

    [XlsTranslate(DictPrefix = "eshop.categories.")]    // Custom value translation
    public string CategoryName { get; set; }

    [XlsHeader(Name = "AlternativeHeader")]             // Renaming the header key
    public DateTime CreatedAt { get; set; }

    public Guid SomeExternalId { get; set; }
}

```

> - If DictionaryPrefix is set, all fields (including sheet name) will be referenced as `my.dictionary.section.<colName>`, eg.:
>   - `my.dictionary.section.Code`. 
>   - `my.dictionary.section.OrderSheetName`.
>   - `my.dictionary.section.AlternativeHeader`.
> - DictionaryPrefix does not affect the `XlsTranslate` fields.


## Custom type mapping

You can add any custom or existing type converter with global scope.

**1. Let's have a custom model:**

```cs
public class Driver 
{
    public string Name { get; set; }
    public string Phone { get; set; }
}
```


**2. Create a converter:**

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

**3. Register your converter once:**


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

> ... or you can import it as RawTable (basically a List of object[] for each row).

```cs

public void LoadXls()
{
    var orders = Importer.Open("eshop_orders.xlsx")
        .ImportAsRaw(
            1,                                  // Index of the sheet based on 1. Optional.
            settings => {
                settings.BreakOnError = true    // Throws exception if some value fails to load,
                settings.HasHeader = true       // If the table has a header to be taken in account
            }
        );
}

```