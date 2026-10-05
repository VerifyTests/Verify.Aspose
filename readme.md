# <img src="/src/icon.png" height="30px"> Verify.Aspose

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://github.com/VerifyTests/Verify.Aspose/actions/workflows/build.yml/badge.svg)](https://github.com/VerifyTests/Verify.Aspose/actions/workflows/build.yml)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Aspose.svg)](https://www.nuget.org/packages/Verify.Aspose/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of documents via [Aspose](https://www.aspose.com/).<!-- singleLineInclude: intro. path: /docs/intro.include.md -->

Verifying a document (pdf, docx, xlsx, or pptx) produces:

 * A `.verified.txt` info file with what Aspose reports of the document (its properties and fonts), the count of its pages, and the text of each page: read as markdown for a pdf or a Word document, and as plain text for the slides of a presentation.
 * The document itself as a `.verified.pdf`, `.verified.docx`, `.verified.xlsx` or `.verified.pptx`. It can be left out with [`ExcludeTargets`](#exclude-the-document).
 * A png of every page of a pdf or a Word document, every slide of a presentation, and every sheet of a workbook, as `#page_0001.verified.png`, `#page_0002.verified.png`, etc.
 * A csv of every sheet of a workbook, named by the sheet: `#Sheet1.verified.csv`.

The page files are named, and the text placed, by Verify's [paged documents](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md) support, which every Verify plugin that splits a document into pages shares. So do the settings that [choose what is verified](#choosing-what-is-verified).

**See [Milestones](../../milestones?state=closed) for release notes.**

An [Aspose License](https://purchase.aspose.com/policies/license-types) is required to use this tool.


## Sponsors


### Entity Framework Extensions<!-- include: sponsors. path: /docs/sponsors.include.md -->

[Entity Framework Extensions](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Aspose) is a major sponsor and is proud to contribute to the development this project.

[![Entity Framework Extensions](https://raw.githubusercontent.com/VerifyTests/Verify.Aspose/refs/heads/main/docs/zzz.png)](https://entityframework-extensions.net/?utm_source=simoncropp&utm_medium=Verify.Aspose)

### Developed using JetBrains IDEs

[![JetBrains logo.](https://raw.githubusercontent.com/VerifyTests/Verify.Aspose/main/docs/jetbrains.png)](https://jb.gg/OpenSourceSupport)<!-- endInclude -->


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


### Choosing what is verified

What a document is split into is controlled by Verify's settings for [paged documents](https://github.com/VerifyTests/Verify/blob/main/docs/paged-documents.md). Anything left out is not produced at all (pages are not drawn, text is not read, sheets are not exported), so these also save work.

The text of a pdf or a Word document is read as markdown, page by page, and that of a presentation as plain text, slide by slide. It is in the info file by default, under the page it is on. `PageText` moves it to a file for each page (`#page_0001.verified.md`, or `#page_0001.verified.txt` for a slide), or leaves it out with `PageTextPlacement.None`:

<!-- snippet: PageTextPerPage -->
<a id='snippet-PageTextPerPage'></a>
```cs
[Test]
public Task PageTextPerPage() =>
    VerifyFile("sample.docx")
        .PageText(PageTextPlacement.PerPage)
        .ExcludeDerivedTargets("png");
```
<sup><a href='/src/Tests/Samples.cs#L309-L317' title='Snippet source file'>snippet source</a> | <a href='#snippet-PageTextPerPage' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

`PagesToInclude` limits the pages that are drawn and read, to the first pages of a document or to those a delegate accepts. A slide of a presentation is a page, and so is a sheet of a workbook. The document itself is still verified whole, and `PageCount` in the info file is still the number of pages the document has:

<!-- snippet: PagesToInclude -->
<a id='snippet-PagesToInclude'></a>
```cs
[Test]
public Task PageCountIsIndependentOfPagesToInclude()
{
    var document = BuildThreePageDocument();
    return Verify(document)
        .PagesToInclude(1);
}
```
<sup><a href='/src/Tests/WordPageCountTests.cs#L7-L17' title='Snippet source file'>snippet source</a> | <a href='#snippet-PagesToInclude' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

A sheet that is not included is left out altogether: its png, its csv, and what the info file says of it.

`ExcludeDerivedTargets("png")` leaves out the drawn pages, keeping the document and its text. `ExcludeDerivedTargets("csv")` does the same for the csv of each sheet:

<!-- snippet: TextOnly -->
<a id='snippet-TextOnly'></a>
```cs
[Test]
public Task TextOnly() =>
    VerifyFile("sample.pdf")
        .ExcludeDerivedTargets("png");
```
<sup><a href='/src/Tests/Samples.cs#L319-L326' title='Snippet source file'>snippet source</a> | <a href='#snippet-TextOnly' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

Each can also be set for every test, on `VerifierSettings`:

<!-- snippet: InitializeOutputs -->
<a id='snippet-InitializeOutputs'></a>
```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyAspose.Initialize();

    // For every test: no page is drawn and no sheet is exported to csv,
    // so only the documents and their text are verified
    VerifierSettings.ExcludeDerivedTargets("png", "csv");
}
```
<sup><a href='/src/StaticSettingsTests/ModuleInitializer.cs#L3-L15' title='Snippet source file'>snippet source</a> | <a href='#snippet-InitializeOutputs' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L4-L10' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPdf' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L24-L33' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPdfStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPdf.verified.txt -->
<a id='snippet-Samples.VerifyPdf.verified.txt'></a>
```txt
{
  Document: {
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
    ]
  },
  PageCount: 2,
  Pages: [
    {
      Number: 1,
      Text:
![](content.001.png)

**Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License [**https://products.aspose.com/words/temporary-license/**](https://products.aspose.com/words/temporary-license/)**



<a name="br1"></a>aluation Only. Created with Aspose.PDF. Copyright 2002-2026 Aspose Pty Ltd.

Evaluation Only. Created with Aspose.PDF. Copyright 2002-2026 Aspose Pty Ltd.

A Simple PDF File

This is a small demonstration .pdf file -

just for use in the Virtual Mechanics tutorials. More text. And more

text. And more text. And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. And more text. Boring, zzzzz. And more text. And more text. And

more text. And more text. And more text. And more text. And more text.

And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. And more text. And more text. Even more. Continued on page 2 ...


**Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd.**

    },
    {
      Number: 2,
      Text:
![](content.001.png)

**Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License [**https://products.aspose.com/words/temporary-license/**](https://products.aspose.com/words/temporary-license/)**



<a name="br1"></a>Simple PDF File 2

...continued from page 1. Yet more text. And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. Oh, how boring typing this stuff. But not as boring as watching

paint dry. And more text. And more text. And more text. And more text.

Boring. More, a little more text. The end, and just as well.


**Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd.**

    }
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyPdf.verified.txt#L1-L94' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPdf.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPdf#page_0001.verified.png](/src/Tests/Samples.VerifyPdf%23page_0001.verified.png):

<img src="/src/Tests/Samples.VerifyPdf%23page_0001.verified.png" width="200px">


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-VerifyExcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L66-L72' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcel' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L182-L191' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyExcelStream' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L133-L155' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWorkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

A sheet is a page. `Number` is its position in the workbook, and `Info` has its name, its columns, its custom properties and its hyperlinks. A sheet with nothing in it has no png. A `Worksheet` verified on its own is the one page.

<!-- snippet: Samples.VerifyWorkbook.verified.txt -->
<a id='snippet-Samples.VerifyWorkbook.verified.txt'></a>
```txt
{
  Document: {
    HasMacro: False,
    HasRevisions: False,
    IsDigitallySigned: False,
    Properties: {
      Comments: the comments
    },
    CustomProperties: {
      key: value
    },
    Fonts: [
      Arial
    ]
  },
  PageCount: 3,
  Pages: [
    {
      Number: 1,
      Info: {
        Name: Sheet1
      }
    },
    {
      Number: 2,
      Info: {
        Name: New Sheet,
        Columns: [
          {
            Name: Some Text
          }
        ]
      }
    },
    {
      Number: 3,
      Info: {
        Name: Evaluation Warning,
        Columns: [
          {
            Name: Evaluation Only. Created with Aspose.Cells for .NET. Copyright 2003 - 2026 Aspose Pty Ltd.
          }
        ]
      }
    }
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyWorkbook.verified.txt#L1-L47' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWorkbook.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The `Evaluation Warning` sheet is added by Aspose.Cells when it saves a workbook without a license.

[Samples.VerifyExcel#page_0001.verified.png](/src/Tests/Samples.VerifyExcel%23page_0001.verified.png):

<img src="/src/Tests/Samples.VerifyExcel%23page_0001.verified.png" width="200px">


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-VerifyWord'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx");
```
<sup><a href='/src/Tests/Samples.cs#L239-L245' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWord' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L257-L266' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyWordStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyWord.verified.txt -->
<a id='snippet-Samples.VerifyWord.verified.txt'></a>
```txt
{
  Document: {
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
      Pages: 2,
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
    ]
  },
  PageCount: 2,
  Pages: [
    {
      Number: 1,
      Text:
![ref1]

**Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License [**https://products.aspose.com/words/temporary-license/**](https://products.aspose.com/words/temporary-license/)**

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

**Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd.**

2

[ref1]: content.001.png

    },
    {
      Number: 2,
      Text:
![ref1]

**Created with an evaluation copy of Aspose.Words. To remove all limitations, you can use Free Temporary License [**https://products.aspose.com/words/temporary-license/**](https://products.aspose.com/words/temporary-license/)**

|<h1>Action items</h1>|<h1>Owner(s)</h1>|<h1>Deadline</h1>|<h1>Status</h1>|
| :- | :- | :- | :- |
|[Action item 6]|[Name(s) 6]|[Date 6]|[Status 6]|

**Evaluation Only. Created with Aspose.Words. Copyright 2003-2026 Aspose Pty Ltd.**

2

[ref1]: content.001.png

    }
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyWord.verified.txt#L1-L90' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyWord#page_0001.verified.png](/src/Tests/Samples.VerifyWord%23page_0001.verified.png):

<img src="/src/Tests/Samples.VerifyWord%23page_0001.verified.png" width="200px">


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerPoint -->
<a id='snippet-VerifyPowerPoint'></a>
```cs
[Test]
public Task VerifyPowerPoint() =>
    VerifyFile("sample.pptx");
```
<sup><a href='/src/Tests/Samples.cs#L37-L43' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerPoint' title='Start of snippet'>anchor</a></sup>
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
<sup><a href='/src/Tests/Samples.cs#L45-L54' title='Snippet source file'>snippet source</a> | <a href='#snippet-VerifyPowerPointStream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPowerPoint.verified.txt -->
<a id='snippet-Samples.VerifyPowerPoint.verified.txt'></a>
```txt
{
  Document: {
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
  },
  PageCount: 3,
  Pages: [
    {
      Number: 1,
      Text:
Lorem... text has been truncated due to evaluation version limitation.
Lorem... text has been truncated due to evaluation version limitation.
Maece... text has been truncated due to evaluation version limitation.

    },
    {
      Number: 2,
      Text:
Chart

    },
    {
      Number: 3,
      Text:
Table
Colum... text has been truncated due to evaluation version limitation.
Colum... text has been truncated due to evaluation version limitation.
Colum... text has been truncated due to evaluation version limitation.
Colum... text has been truncated due to evaluation version limitation.
Colum... text has been truncated due to evaluation version limitation.

    }
  ]
}
```
<sup><a href='/src/Tests/Samples.VerifyPowerPoint.verified.txt#L1-L99' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPowerPoint.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPowerPoint#page_0001.verified.png](/src/Tests/Samples.VerifyPowerPoint%23page_0001.verified.png):

<img src="/src/Tests/Samples.VerifyPowerPoint%23page_0001.verified.png" width="200px">


### Binary output across .NET frameworks

When verifying binary package output (xlsx, docx, nupkg, etc.) across multiple target frameworks (e.g. net48 and net10.0), the binary output may differ due to Deflate compression implementation differences. The XML content within entries is identical — only the compressed bytes differ. Use `UniqueForRuntime` to generate framework-specific verified files:

```cs
await Verify(stream, extension: "xlsx")
    .UniqueForRuntime();
```

See [Verify Naming docs](https://github.com/VerifyTests/Verify/blob/main/docs/naming.md) for more details.


## Exclude the document

The source document is included in the snapshot as a `.verified.pdf`, `.verified.docx` or `.verified.xlsx`. Building the deterministic document is expensive, and committing it is not always wanted. [`ExcludeTargets`](https://github.com/VerifyTests/Verify/blob/main/docs/converter.md#excluding-targets) drops it from a verification and skips the build, while the info, text, csv, and rendered pages still verify:

<!-- snippet: ExcludeXlsx -->
<a id='snippet-ExcludeXlsx'></a>
```cs
[Test]
public Task ExcludeXlsx() =>
    // ExcludeTargets skips the expensive deterministic xlsx build.
    VerifyFile("sample.xlsx")
        .ExcludeTargets("xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L74-L82' title='Snippet source file'>snippet source</a> | <a href='#snippet-ExcludeXlsx' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

The same applies to `docx` via `ExcludeTargets("docx")` and to `pdf` via `ExcludeTargets("pdf")`. A `doc` is given back as a `docx`, and an `xls` as an `xlsx`, so those are the extensions to exclude for them. To exclude for every test, call `VerifierSettings.ExcludeTargets("xlsx")` at initialization.


## Reviewing changes

A change to a document is a change to several files: the document, its info file, and every page. Verify tells the diff tool that the pages, the csv of each sheet and the info file were derived from the document, and [DiffEngineViewer](https://github.com/VerifyTests/DiffEngine/blob/main/docs/viewer.md#files-derived-from-a-document), which draws the pages of a document itself, shows them as one row and accepts them together. Other diff tools are given each file, as before.

When the document differs, its pages are compared exactly, skipping any [comparer](https://github.com/VerifyTests/Verify/blob/main/docs/comparer.md) registered for png.


## Migrating from 5.x

Version 6 moves to the paged document support in Verify 33.3. The `outputs` argument of `Initialize` is gone, as is the `PagesToInclude` of this package, and the `AsposeOutputs` enum is obsolete as an error, so that code naming it is pointed here. Verify's settings replace them, and can be set for one verification as well as for every test:

| 5.x | 6.x |
| --- | --- |
| `Initialize` without `AsposeOutputs.Png` | `ExcludeDerivedTargets("png")` |
| `Initialize` without `AsposeOutputs.Text` | `PageText(PageTextPlacement.None)` |
| `Initialize` without `AsposeOutputs.Csv` | `ExcludeDerivedTargets("csv")` |
| `Initialize(AsposeOutputs.Text)` | `VerifierSettings.ExcludeDerivedTargets("png", "csv")` |
| `Initialize(AsposeOutputs.None)` | `VerifierSettings.PageText(PageTextPlacement.None)` and `VerifierSettings.ExcludeDerivedTargets("png", "csv")` |
| `.PagesToInclude(count)` from `VerifyTestsAspose` | `.PagesToInclude(count)` from Verify. The call is the same |

`PagesToInclude` now also limits the sheets of a workbook, where it was ignored.

The snapshot files are renamed:

| 5.x | 6.x |
| --- | --- |
| `Tests.Pdf#00.verified.png`, `Tests.Pdf#01.verified.png` | `Tests.Pdf#page_0001.verified.png`, `Tests.Pdf#page_0002.verified.png` |
| `Tests.Pdf.verified.png`, where there was one page | `Tests.Pdf#page_0001.verified.png` |
| `Tests.Excel#Sheet1.verified.png` | `Tests.Excel#page_0001.verified.png` |
| `Tests.Word.verified.xml`, from `IncludeWordStyles` | `Tests.Word#styles.verified.xml` |
| `Tests.Excel#Sheet1.verified.csv` | Unchanged |
| `Tests.Word.verified.docx`, and the pdf and xlsx | Unchanged |
| A presentation was not a snapshot | `Tests.PowerPoint.verified.pptx` |
| `Tests.Word.verified.txt` | Same name, new shape |

A presentation is now a snapshot, as the other documents are. It is saved as a pptx, a `ppt` included, and passed through [DeterministicIoPackaging](https://github.com/SimonCropp/DeterministicIoPackaging), with the field ids and the last printed time that Aspose.Slides writes afresh on each save neutralized. `VerifierSettings.ExcludeTargets("pptx")` keeps it out of the snapshots, as it was in 5.x.

Renamed snapshots show as a new file and a pending delete. Accepting both, or running once with [AutoVerify](https://github.com/VerifyTests/Verify/blob/main/docs/autoverify.md), moves a test over. The content of a png is unchanged, so source control shows it as a rename.

The info file has the shape every paged document has. What Aspose reports of the document is under `Document`, `PageCount` is the number of pages, slides or sheets (it was `Pages` for a pdf), and `Text` follows:

```
{                                    {
  PageCount: 2,                        Document: {
  HasRevisions: False,                   HasRevisions: False,
  Fonts: [                               Fonts: [
    Arial                                  Arial
  ],                                     ]
  Text: The text                       },
}                                      PageCount: 2,
                                       Text: The text
                                     }
```

For a workbook, what was in `Sheets` is the `Info` of each page. For a `Worksheet` verified on its own, what was the whole info file is the `Info` of the one page:

```
{                                    {
  HasMacro: False,                     Document: {
  Sheets: [                              HasMacro: False
    {                                  },
      Name: Sheet1                     PageCount: 1,
    }                                  Pages: [
  ]                                      {
}                                          Number: 1,
                                           Info: {
                                             Name: Sheet1
                                           }
                                         }
                                       ]
                                     }
```

A document that is itself a named target of a verification, an attachment for example, has its files named by Verify: `#Attachment1.Sheet1.verified.csv`, where it was `#Attachment1-Sheet1.verified.csv`.


## File Samples

http://file-examples.com/


## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/).
