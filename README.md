# SheetsPersist
Simple Object Persistence and Logging to Google Sheets

SheetsPersist gives you two ways to integrate .NET object instances with Google Sheets.

1. Object Persistence (read and write object instances to/from a spreadsheet).
2. Logging (send .NET instances to a rich log kept in a spreadsheet).

## Spreadsheet Registration
First, register the spreadsheet you want to work with like this:

```csharp
GoogleSheets.RegisterDocumentID("LogTest", "1YnqCNZFqfdRP_7ocmAkD0dpI-G_bFnwGH1vN-YppAZE");
```

The first parameter is the name of the spreadsheet as you would like to refer to it in the code.

The second parameter is the ID of the document. You can get this from the URL.

For example:

`docs.google.com/spreadsheets/d/1YnqCNZFqfdRP_7ocmAkD0dpI-G_bFnwGH1vN-YppAZE/edit#gid=835270728`

Registration needs to happen when the app first starts up (before you write any data). 

## Add Attributes to Classes
Any class that needs to be written to or read from a Google Sheet, needs to be adorned with two attributes:

The Document attribute includes the name of the document passed to RegisterDocumentID (see above). This is the 
spreadsheet that will hold saved instances.

```csharp
[Document("LogTest")]
```


The Sheet attribute contains the name of the sheet (tab) to read from or write to.

```csharp
[Sheet("Today's Results")]
```

## Object Persistence

Specify the properties and fields to read/write using the column attribute, like this:

```csharp
[Column]
public DateTime Time { get; set; }
```

In this case, the column header will be "Time". If you want to use a different column name, pass that string in as a parameter to the column attribute, like this:

```csharp
[Column("Sale Price")]
public decimal Price { get; set; }
```

You can specify an indexer property (which must have unique values in the spreadsheet), like this:

```csharp
[Indexer]
[Column]
public Guid ID { get; set; }
```

If a class has an indexer, when you save it, it will overwrite the data on row where that class data previously existed.

### Saving

To save objects adorned with the Document and Sheet attributes (see above), simply call:

```csharp
GoogleSheets.SaveChanges(...)
```
Passing the instance or instances you want to save. That's it.

### Loading

To load data from a spreadsheet, use:

```csharp
List<T> results = GoogleSheets.Get<T>()
```
Where T is the class that has been adorned with the Document and Sheet attributes. That's it.

## Object Logging

Object logging is super cool. You can rapidly send many rich log messages to a spreadsheet that you can monitor from anywhere and share easily.

To log an object, call:

```csharp
GoogleSheets.AppendRow(instance);
```

Where *instance* is an instance of the class that has been adorned with the Document and Sheet attributes.

### Throttling

Object logging is throttled, so you can send many messages in a short time without exceeding Google API messaging quotas. 

The default time between updates is 5 seconds. You can change it like this if you like:

```csharp
GoogleSheets.TimeBetweenThrottledUpdates = TimeSpan.FromSeconds(10);
```

We do not recommend dropping this time below 2 seconds because of Google Api Messaging Quotas.

