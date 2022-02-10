# SheetsPersist
Simple Object Persistence and Logging to Google Sheets

SheetsPersist gives you two ways to integrate .NET object instances with Google Sheets.

1. Object Persistence (read and write object instances to/from a spreadsheet).
2. Logging (send .NET instances to a rich log kept in a spreadsheet).

## Spreadsheet Registration
The first thing you need to do is register your spreadsheet. You can do this with a call like this:

```csharp
GoogleSheets.RegisterDocumentID("LogTest", "1YnqCNZFqfdRP_7ocmAkD0dpI-G_bFnwGH1vN-YppAZE");
```

The first parameter is the name of the spreadsheet as you would like to refer to it in the code.

The second parameter is the ID of the document. You can get this from the URL.

For example:

docs.google.com/spreadsheets/d/1YnqCNZFqfdRP_7ocmAkD0dpI-G_bFnwGH1vN-YppAZE/edit#gid=835270728

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
[Sheet("CGR", 1)]
```

## Object Persistence

To persist objects, adorn the class with the 

## Object Logging


