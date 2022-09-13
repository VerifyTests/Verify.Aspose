# <img src="/src/icon.png" height="30px"> Verify.Aspose

[![Build status](https://ci.appveyor.com/api/projects/status/7k8hh0guut2ioak2?svg=true)](https://ci.appveyor.com/project/SimonCropp/Verify-Aspose)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Aspose.svg)](https://www.nuget.org/packages/Verify.Aspose/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of documents via [Aspose](https://www.aspose.com/).

Converts documents (pdf, docx, xslx, and pptx) to png for verification.

An [Aspose License](https://purchase.aspose.com/policies/license-types) is required to use this tool.



## NuGet package

https://nuget.org/packages/Verify.Aspose/


## Usage


### Enable Verify.Aspose

<!-- snippet: enable -->
<a id='snippet-enable'></a>
```cs
[ModuleInitializer]
public static void Initialize()
{
    VerifyAspose.Initialize();
```
<sup><a href='/src/Tests/ModuleInitializer.cs#L3-L8' title='Snippet source file'>snippet source</a> | <a href='#snippet-enable' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


### PDF


#### Verify a file

<!-- snippet: VerifyPdf -->
<a id='snippet-verifypdf'></a>
```cs
[Test]
public Task VerifyPdf() =>
    VerifyFile("sample.pdf");
```
<sup><a href='/src/Tests/Samples.cs#L8-L14' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdf' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPdfStream -->
<a id='snippet-verifypdfstream'></a>
```cs
[Test]
public Task VerifyPdfStream() =>
    Verify(File.OpenRead("sample.pdf"))
        .UseExtension("pdf");
```
<sup><a href='/src/Tests/Samples.cs#L28-L35' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdfstream' title='Start of snippet'>anchor</a></sup>
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
  Version: 1.4
}
```
<sup><a href='/src/Tests/Samples.VerifyPdf.verified.txt#L1-L25' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPdf.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPdf.01.verified.png](/src/Tests/Samples.VerifyPdf.01.verified.png):

<img src="/src/Tests/Samples.VerifyPdf.01.verified.png" width="200px">


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-verifyexcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L58-L64' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-verifyexcelstream'></a>
```cs
[Test]
public Task VerifyExcelStream() =>
    Verify(File.OpenRead("sample.xlsx"))
        .UseExtension("xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L101-L108' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcelstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a WorkBook

<!-- snippet: VerifyWorkbook -->
<a id='snippet-verifyworkbook'></a>
```cs
[Test]
public Task VerifyWorkbook()
{
    var book = new Workbook();

    var sheet = book.Worksheets.Add("New Sheet");

    var cells = sheet.Cells;

    cells[0, 0].PutValue("Some Text");
    return Verify(book);
}
```
<sup><a href='/src/Tests/Samples.cs#L84-L99' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyworkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyExcel.01.verified.png](/src/Tests/Samples.VerifyExcel.01.verified.png):

<img src="/src/Tests/Samples.VerifyExcel.01.verified.png" width="200px">


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-verifyword'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx");
```
<sup><a href='/src/Tests/Samples.cs#L110-L116' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyword' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-verifywordstream'></a>
```cs
[Test]
public Task VerifyWordStream() =>
    Verify(File.OpenRead("sample.docx"))
        .UseExtension("docx");
```
<sup><a href='/src/Tests/Samples.cs#L118-L125' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifywordstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyWord.verified.txt -->
<a id='snippet-Samples.VerifyWord.verified.txt'></a>
```txt
{
  HasRevisions: False,
  DefaultLocale: EnglishUS,
  Properties: {
    CreateTime: DateTime_1,
    Language: en-US,
    TotalEditingTime: 991904
  }
}
```
<sup><a href='/src/Tests/Samples.VerifyWord.verified.txt#L1-L9' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyWord.01.verified.png](/src/Tests/Samples.VerifyWord.01.verified.png):

<img src="/src/Tests/Samples.VerifyWord.01.verified.png" width="200px">


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerPoint -->
<a id='snippet-verifypowerpoint'></a>
```cs
[Test]
public Task VerifyPowerPoint() =>
    VerifyFile("sample.pptx");
```
<sup><a href='/src/Tests/Samples.cs#L39-L45' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypowerpoint' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPowerPointStream -->
<a id='snippet-verifypowerpointstream'></a>
```cs
[Test]
public Task VerifyPowerPointStream() =>
    Verify(File.OpenRead("sample.pptx"))
        .UseExtension("pptx");
```
<sup><a href='/src/Tests/Samples.cs#L47-L54' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypowerpointstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPowerPoint.00.verified.txt -->
<a id='snippet-Samples.VerifyPowerPoint.00.verified.txt'></a>
```txt
{
  AppVersion: 16.0000,
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
  HyperlinkBase: 
}
```
<sup><a href='/src/Tests/Samples.VerifyPowerPoint.00.verified.txt#L1-L23' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPowerPoint.00.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPowerPoint.01.verified.png](/src/Tests/Samples.VerifyPowerPoint.01.verified.png):


## File Samples

http://file-examples.com/



## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/).
