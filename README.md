# SheetsPersist
Simple **Object Persistence** and **Logging** to Google Sheets

SheetsPersist gives you two ways to integrate .NET object instances with Google Sheets:

1. Object Persistence (read and write object instances to/from a spreadsheet).
2. Logging (send .NET instances to a rich log kept in a spreadsheet).

## Spreadsheet Registration
First, register the spreadsheet you want to work with like this:

```csharp
GoogleSheets.RegisterDocumentID("LogTest", "1YnqCNZFqfdRP_7ocmAkD0dpI-G_bFnwGH1vN-YppAZE");
```

The first parameter is the name of the spreadsheet (as you would like to refer to it in the code).

The second parameter is the **ID** of the document. You can get this from the Google Sheets URL.

For example:

![image](https://user-images.githubusercontent.com/4841528/159050489-ea17bdcd-62d9-4405-801d-9b50fc146e97.png)

Registration needs to happen before the app attempts to read or write any data. A good place to make this registration call is from a **static constructor** in the class you are persisting.

## Add Attributes to Classes
Serialized classes need two attributes:

The **Document** attribute includes the name of the document passed to RegisterDocumentID (see above). This is the 
spreadsheet that will hold saved instances.

```csharp
[Document("LogTest")]
```


The **Sheet** attribute contains the name of the sheet (tab) to read from or write to. 

```csharp
[Sheet("Today's Results")]
```
Sheets are automatically created if needed.

## Object Persistence

Specify the properties and fields to read/write using the **Column** attribute, like this:

```csharp
[Column]
public DateTime Time { get; set; }
```

In the example above, the column header will be "Time". If you want to use a different column name (something other than the property name), specify that column name as a parameter to the column attribute, like this:

```csharp
[Column("Sale Price")]
public decimal Price { get; set; }
```

You can specify an indexer property for a column containing unique values in the spreadsheet (like a GUID), using the **Indexer** attribute:

```csharp
[Indexer]
[Column]
public Guid ID { get; set; }
```

Note that if  a class has an indexer **and you save it**, it will **overwrite** the data on row in the spreadsheet where the old class data previously existed.

### Saving

To save objects adorned with the Document and Sheet attributes (see above), simply call:

```csharp
GoogleSheets.SaveChanges(...)
```
Passing the instance or instances you want to save. That's it.

### Loading

To load all the data for a particular class from a spreadsheet, use:

```csharp
List<T> results = GoogleSheets.Get<T>()
```
Where T is the class that has been adorned with the Document and Sheet attributes. That's it.

## Object Logging

Object logging is super cool. You can rapidly send many rich log messages to a spreadsheet (that you can monitor from anywhere and share easily).

To log an object, call:

```csharp
GoogleSheets.AppendRow(instance);
```

Where *instance* is an instance of the class that has been adorned with the Document and Sheet attributes.

Logging to spreadsheets is also cool because you can use Google Sheet's conditional formatting to emphasize important data (like error messages). And you can store multi-line string data to a single cell (for example, a call stack).

### Custom Sheet Names
The logging API lets you to optionally override the sheet name when you log instances. So for example, if you need to organize your logs by month, you can do something like this:

```csharp
static string GetMonthlySheetName()
{
    return DateTime.UtcNow.ToString("MMM yyyy");
}

public static void Log(this MyBuyOrderDto buyOrder)
{
    GoogleSheets.AppendRow(buyOrder, GetMonthlySheetName());
}
```

And your logs will simply appear on the sheet name you specify:

![image](https://user-images.githubusercontent.com/4841528/159039561-2196ddce-824b-42e8-be3c-ca37124f766c.png)


### Additional Attributes

For a richer logging experience, you can add optional attributes.

#### HeaderRow

The **HeaderRow** attribute goes on the class you're logging and allows you to specify a **font color** (use an HTML color string) and an optional **font weight** for the header row (used when creating new sheets).

```csharp
	[Document("Alerts")]
	[Sheet("Active")]
	[HeaderRow("#3d79d7", FontWeight.Bold)]
	public class Alert
  ...
```

#### FormatDate
The **FormatDate** attribute lets you specify a pattern for formatting dates, like this:

```csharp
[Column("Time")]
[FormatDate("dd MMM yyyy - HH:mm:ss")]
public DateTime OrderDateTime { get; set; }
```

![image](https://user-images.githubusercontent.com/4841528/159042750-dfcf2e1c-27b6-456e-85d6-3f9f4589a877.png)

#### FormatCurrency

The **FormatCurrency** attribute lets you specify a pattern for formatting currency, like this:

```csharp
[Column("Transfered")]
[FormatCurrency("\\$#,##0.00;[red](\\$#,##0.00)")]
public decimal TotalTransfered { get; set; }
```

![image](https://user-images.githubusercontent.com/4841528/159046218-675cf1ba-a21d-4c2f-ba83-8fb1038be640.png)

See [this link](https://www.benlcollins.com/spreadsheets/google-sheets-custom-number-format/) for a useful guide on formatting in Google Sheets.

#### FormatNumber

The **FormatNumber** attribute lets you specify a pattern for formatting numbers, like this:

```csharp
[Column("% Change")]
[FormatNumber("0.##\"%\"")]
public double PercentChange { get; set; }
```

![image](https://user-images.githubusercontent.com/4841528/159047085-4100eee2-22ac-460c-92c7-bbd38ee7ba3a.png)


#### Note
The **Note** attribute lets you specify a note to explain the contents of the column, like this:

```csharp
[Column("Remaining")]
[Note("The unfilled amount of the order remaining (original order quantity - order filled).")]
public decimal QuantityRemaining => quantityRemaining;
```

The note appears when you hover the mouse over the column header.

![image](https://user-images.githubusercontent.com/4841528/159047979-eebf6022-ef17-430d-8e53-75768c6d49e3.png)

#### Freezing Rows and Columns

The **Sheet** attribute for the class also lets you specify the number of rows and/or columns you would like to freeze (so they stay visible as you scroll through the data). Just add these arguments to the attribute if you need them. For example, to freeze the top row, pass a "1" into the frozenRowCount parameter, like this:

```csharp
[Sheet("Alerts", 1)]
```

### Throttling

Object logging is throttled, so you can send many messages in a short time without exceeding Google API messaging quotas. 

The default time between updates is 5 seconds. You can change it like this if you like:

```csharp
GoogleSheets.TimeBetweenThrottledUpdates = TimeSpan.FromSeconds(10);
```

We recommend keeping this time above 2 seconds to stay below Google Api Messaging Quotas.

### Custom Read/Write Formats

You can specify data formats for values written to and read from Google Sheets. Just add one or both of the optional enums to the [Sheet] attribute, like this:

[Sheet("Data", 1, 0, ReadValuesAs.Unformatted, WriteValuesAs.Raw)]

Available read options are:

* **Formatted** - (**default read format if not specified**) Values will be calculated & formatted in the reply according to the cell's formatting. Formatting is based on the spreadsheet's locale, not the requesting user's locale. For example, if `A1` is `1.23` and `A2` is `=A1` and formatted as currency, then `A2` would return `"$1.23"`.
* **Unformatted** - Values will be calculated, but not formatted in the reply. For example, if `A1` is `1.23` and `A2` is `=A1` and formatted as currency, then `A2` would return the number `1.23`.
* **Formula** - Values will not be calculated. The reply will include the formulas. For example, if `A1` is `1.23` and `A2` is `=A1` and formatted as currency, then A2 would return `"=A1"`.

Available write options are:
* **Raw** - The values transferred to the Sheet will not be parsed and will be stored as-is.
* **UserEntered** - (**default write format if not specified**) The values will be parsed as if the user typed them into the UI. Numbers will stay as numbers, but strings may be converted to numbers, dates, etc., following the same rules normally applied when entering text into a cell via the Google Sheets UI.

## Updates
1.3.4 - Added support for specifying formats for reading and writing. Added support for the Int64 (long) property type. Improved error message content.

1.3.3 - Added concurrency lock around the MessageThrottlers and removed calls to System.Debugger.Break().
