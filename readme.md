# <img src="/src/icon.png" height="30px"> Verify.Aspose

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://img.shields.io/appveyor/build/SimonCropp/Verify-Aspose)](https://ci.appveyor.com/project/SimonCropp/Verify-Aspose)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Aspose.svg)](https://www.nuget.org/packages/Verify.Aspose/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of documents via [Aspose](https://www.aspose.com/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Converts documents (pdf, docx, xlsx, and pptx) to png for verification.

**See [Milestones](../../milestones?state=closed) for release notes.**

An [Aspose License](https://purchase.aspose.com/policies/license-types) is required to use this tool.


## Sponsors


### Entity Framework Extensions<!-- include: zzz. path: /docs/zzz.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Aspose) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.Aspose/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Aspose)<!-- endInclude -->


## NuGet

 * https://nuget.org/packages/Verify.Aspose


## Usage


### Enable Verify.Aspose

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize() =>
    VerifyAspose.Initialize();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PDF


#### Verify a file

<!-- snippet: VerifyPdf -->
<a id='snippet-VerifyPdf'></a>
```cs
[Test]
public Task VerifyPdf() =>
    VerifyFile("sample.pdf");
```
<sup><a href='/src/Tests/Samples.cs#L5-L11' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPdf' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPdfStream -->
<a id='snippet-VerifyPdfStream'></a>
```cs
[Test]
public Task VerifyPdfStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pdf"));
    return Verify(stream, "pdf");
}
```
<sup><a href='/src/Tests/Samples.cs#L25-L34' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPdfStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPdf.verified.txt -->
<a id='snippet-Samples.VerifyPdf.verified.txt'></a>
```txt
{
  Pages: 2,
  AllowReusePageContent: false,
  CenterWindow: false,
  DisplayDocTitle: false,
  FitWindow: False,
  HideMenubar: False,
  HideToolBar: False,
  HideWindowUI: False,
  IgnoreCorruptedObjects: True,
  Info: {
    Creator: RAD PDF,
    Producer: RAD PDF 3.9.0.0 - http://www.radpdf.com
  },
  IsEncrypted: False,
  IsLinearized: False,
  IsPdfaCompliant: False,
  IsPdfUaCompliant: False,
  IsXrefGapsAllowed: True,
  OptimizeSize: False,
  PageLabels: {},
  PageLayout: Default,
  PdfFormat: v_1_4,
  Version: 1.4,
  Fonts: [
    Helvetica
  ],
  Text:

<a name="br1"></a>A Simple PDF File

This is a small demonstration .pdf file -

just for use in the Virtual Mechanics tutorials. More text. And more

text. And more text. And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. And more text. Boring, zzzzz. And more text. And more text. And

more text. And more text. And more text. And more text. And more text.

And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. And more text. And more text. Even more. Continued on page 2 ...





<a name="br2"></a> 

Simple PDF File 2

...continued from page 1. Yet more text. And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. Oh, how boring typing this stuff. But not as boring as watching

paint dry. And more text. And more text. And more text. And more text.

Boring. More, a little more text. The end, and just as well.



}
```
<sup><a href='/src/Tests/Samples.VerifyPdf.verified.txt#L1-L70' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPdf.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPdf#00.verified.png](/src/Tests/Samples.VerifyPdf%2300.verified.png):

<img src="/src/Tests/Samples.VerifyPdf%2300.verified.png" width="200px">


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L67-L73' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-VerifyExcelStream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(stream, "xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L143-L152' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a WorkBook

<!-- snippet: VerifyWorkbook -->
<a id='snippet-VerifyWorkbook'></a>
```cs
[Test]
public Task VerifyWorkbook()
{
    var book = new Workbook
    {
        BuiltInDocumentProperties =
        {
            Comments = "the comments"
        }
    };
    book.CustomDocumentProperties.Add("key", "value");

    var sheet = book.Worksheets.Add("New Sheet");

    var cells = sheet.Cells;

    cells[0, 0].PutValue("Some Text");
    return Verify(book);
}
```
<sup><a href='/src/Tests/Samples.cs#L108-L130' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWorkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyExcel#Sheet1.verified.png](/src/Tests/Samples.VerifyExcel#Sheet1.verified.png):

<img src="/src/Tests/Samples.VerifyExcel%23Sheet1.verified.png" width="200px">


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-VerifyWord'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx");
```
<sup><a href='/src/Tests/Samples.cs#L201-L207' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWord' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-VerifyWordStream'></a>
```cs
[Test]
public Task VerifyWordStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
    return Verify(stream, "docx");
}
```
<sup><a href='/src/Tests/Samples.cs#L209-L218' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWordStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyWord.verified.txt -->
<a id='snippet-Samples.VerifyWord.verified.txt'></a>
```txt
{
  HasRevisions: False,
  DefaultLocale: EnglishUS,
  Properties: {
    Characters: 1009,
    CharactersWithSpaces: 1183,
    CreateTime: DateTime_1,
    HeadingPairs: [
      Title,
      1
    ],
    LastSavedTime: DateTime_2,
    Lines: 8,
    Pages: 1,
    Paragraphs: 2,
    Template: Normal,
    Words: 176
  },
  CustomProperties: {
    ContentTypeId: 0x010100AA3F7D94069FF64A86F7DFF56D60E3BE
  },
  ShadeFormData: true,
  Fonts: [
    Consolas,
    Segoe UI,
    Symbol,
    Times New Roman,
    Trebuchet MS
  ],
  Text:
[Meeting name] meeting minutes

|Location:|[Address or room number]|
| :- | :- |
|Date:|[Date]|
|Time:|[Time]|
|Attendees:|[List attendees]|
# Agenda items
1. [It’s easy to make this template your own. To replace placeholder text, just select it and start typing. Don’t include space to the right or left of the characters in your selection.]
1. [Apply any text formatting you see in this template with just a click from the Home tab, in the Styles group. For example, this text uses the List Number style.]
1. [To add a new row at the end of the action items table, just click into the last cell in the last row and then press Tab.]
1. [To add a new row or column anywhere in a table, click in an adjacent row or column to the one you need and then, on the Table Tools Layout tab of the ribbon, click an Insert option.]
1. [Agenda item]
1. [Agenda item]

|<h1>Action items</h1>|<h1>Owner(s)</h1>|<h1>Deadline</h1>|<h1>Status</h1>|
| :- | :- | :- | :- |
|[Action item 1]|[Name(s) 1]|[Date 1]|[Status 1, such as In Progress or Complete]|
|[Action item 2]|[Name(s) 2]|[Date 2]|[Status 2]|
|[Action item 3]|[Name(s) 3]|[Date 3]|[Status 3]|
|[Action item 4]|[Name(s) 4]|[Date 4]|[Status 4]|
|[Action item 5]|[Name(s) 5]|[Date 5]|[Status 5]|
|[Action item 6]|[Name(s) 6]|[Date 6]|[Status 6]|

2

}
```
<sup><a href='/src/Tests/Samples.VerifyWord.verified.txt#L1-L57' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyWord#00.verified.png](/src/Tests/Samples.VerifyWord%2300.verified.png):

<img src="/src/Tests/Samples.VerifyWord%2300.verified.png" width="200px">


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerPoint -->
<a id='snippet-VerifyPowerPoint'></a>
```cs
[Test]
public Task VerifyPowerPoint() =>
    VerifyFile("sample.pptx");
```
<sup><a href='/src/Tests/Samples.cs#L38-L44' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerPoint' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPowerPointStream -->
<a id='snippet-VerifyPowerPointStream'></a>
```cs
[Test]
public Task VerifyPowerPointStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
    return Verify(stream, "pptx");
}
```
<sup><a href='/src/Tests/Samples.cs#L46-L55' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerPointStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPowerPoint.verified.txt -->
<a id='snippet-Samples.VerifyPowerPoint.verified.txt'></a>
```txt
{
  Properties: {
    NameOfApplication: Microsoft Office PowerPoint,
    Company: ,
    Manager: ,
    PresentationFormat: Custom,
    SharedDoc: false,
    ApplicationTemplate: ,
    Title: Lorem ipsum,
    Subject: ,
    Author: simon,
    Keywords: ,
    Comments: ,
    Category: ,
    CreatedTime: DateTime_1,
    LastSavedTime: DateTime_2,
    LastPrinted: DateTime_3,
    LastSavedBy: Simon Cropp,
    RevisionNumber: 1,
    ContentStatus: ,
    ContentType: ,
    HyperlinkBase: ,
    ScaleCrop: false,
    LinksUpToDate: false,
    HyperlinksChanged: false,
    Slides: 3,
    Notes: 3,
    Paragraphs: 14,
    Words: 231,
    TitlesOfParts: [
      Times New Roman,
      Arial,
      Droid Sans Fallback,
      WenQuanYi Zen Hei,
      DejaVu Sans,
      Office Theme,
      Office Theme,
      Lorem ipsum,
      Chart,
      Table
    ],
    HeadingPairs: [
      {
        Name: Fonts Used,
        Count: 5
      },
      {
        Name: Theme,
        Count: 2
      },
      {
        Name: Embedded OLE Servers
      },
      {
        Name: Slide Titles,
        Count: 3
      }
    ]
  },
  Fonts: [
    Arial,
    Calibri,
    Calibri Light,
    DejaVu Sans,
    Droid Sans Fallback,
    Times New Roman,
    WenQuanYi Zen Hei
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyPowerPoint.verified.txt#L1-L69' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPowerPoint.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPowerPoint%2300.verified.png](/src/Tests/Samples.VerifyPowerPoint%2300.verified.png):

<img src="/src/Tests/Samples.VerifyPowerPoint%2300.verified.png" width="200px">


## File Samples

http://file-examples.com/


## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/).
