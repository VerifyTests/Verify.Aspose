# <img src="/src/icon.png" height="30px"> Verify.Aspose

[![Discussions](https://img.shields.io/badge/Verify-Discussions-yellow?svg=true&label=)](https://github.com/orgs/VerifyTests/discussions)
[![Build status](https://ci.appveyor.com/api/projects/status/7k8hh0guut2ioak2?svg=true)](https://ci.appveyor.com/project/SimonCropp/Verify-Aspose)
[![NuGet Status](https://img.shields.io/nuget/v/Verify.Aspose.svg)](https://www.nuget.org/packages/Verify.Aspose/)

Extends [Verify](https://github.com/VerifyTests/Verify) to allow verification of documents via [Aspose](https://www.aspose.com/).

Converts documents (pdf, docx, xslx, and pptx) to png for verification.

**See [Milestones](../../milestones?state=closed) for release notes.**

An [Aspose License](https://purchase.aspose.com/policies/license-types) is required to use this tool.


## NuGet package

https://nuget.org/packages/Verify.Aspose/


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
<a id='snippet-verifypdf'></a>
```cs
[Test]
public Task VerifyPdf() =>
    VerifyFile("sample.pdf");
```
<sup><a href='/src/Tests/Samples.cs#L10-L16' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdf' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPdfStream -->
<a id='snippet-verifypdfstream'></a>
```cs
[Test]
public Task VerifyPdfStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pdf"));
    return Verify(stream, "pdf");
}
```
<sup><a href='/src/Tests/Samples.cs#L30-L39' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypdfstream' title='Start of snippet'>anchor</a></sup>
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
  Text:
![ref1]

**Evaluation Only. Created with Aspose.Words. Copyright 2003-2024 Aspose Pty Ltd.**



<a name="br1"></a>Evaluation Only. Created with Aspose.PDF. Copyright 2002-2023 Aspose Pty Ltd.

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


**Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs please visit: https://products.aspose.com/words/**
![ref2]



<a name="br2"></a> 

Simple PDF File 2

...continued from page 1. Yet more text. And more text. And more text.

And more text. And more text. And more text. And more text. And more

text. Oh, how boring typing this stuff. But not as boring as watching

paint dry. And more text. And more text. And more text. And more text.

Boring. More, a little more text. The end, and just as well.


**Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs please visit: https://products.aspose.com/words/**

[ref1]: content.001.png
[ref2]: content.002.png

}
```
<sup><a href='/src/Tests/Samples.VerifyPdf.verified.txt#L1-L80' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPdf.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPdf#00.verified.png](/src/Tests/Samples.VerifyPdf%2300.verified.png):

<img src="/src/Tests/Samples.VerifyPdf%2300.verified.png" width="200px">


### Excel


#### Verify a file

<!-- snippet: VerifyExcel -->
<a id='snippet-verifyexcel'></a>
```cs
[Test]
public Task VerifyExcel() =>
    VerifyFile("sample.xlsx");
```
<sup><a href='/src/Tests/Samples.cs#L71-L77' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcel' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyExcelStream -->
<a id='snippet-verifyexcelstream'></a>
```cs
[Test]
public Task VerifyExcelStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.xlsx"));
    return Verify(stream, "xlsx");
}
```
<sup><a href='/src/Tests/Samples.cs#L132-L141' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyexcelstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a WorkBook

<!-- snippet: VerifyWorkbook -->
<a id='snippet-verifyworkbook'></a>
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
<sup><a href='/src/Tests/Samples.cs#L97-L119' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyworkbook' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyExcel.verified.png](/src/Tests/Samples.VerifyExcel.verified.png):

<img src="/src/Tests/Samples.VerifyExcel.verified.png" width="200px">


### Word


#### Verify a file

<!-- snippet: VerifyWord -->
<a id='snippet-verifyword'></a>
```cs
[Test]
public Task VerifyWord() =>
    VerifyFile("sample.docx");
```
<sup><a href='/src/Tests/Samples.cs#L143-L149' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifyword' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyWordStream -->
<a id='snippet-verifywordstream'></a>
```cs
[Test]
public Task VerifyWordStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.docx"));
    return Verify(stream, "docx");
}
```
<sup><a href='/src/Tests/Samples.cs#L151-L160' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifywordstream' title='Start of snippet'>anchor</a></sup>
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
    Language: en-US
  },
  Text:
![](content.001.png)

**Evaluation Only. Created with Aspose.Words. Copyright 2003-2024 Aspose Pty Ltd.**

**Lorem ipsum** 

# **Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ac faucibus odio.** 

Vestibulum neque massa, scelerisque sit amet ligula eu, congue molestie mi. Praesent ut varius sem. Nullam at porttitor arcu, nec lacinia nisi. Ut ac dolor vitae odio interdum condimentum. **Vivamus dapibus sodales ex, vitae malesuada ipsum cursus convallis. Maecenas sed egestas nulla, ac condimentum orci.** Mauris diam felis, vulputate ac suscipit et, iaculis non est. Curabitur semper arcu ac ligula semper, nec luctus nisl blandit. Integer lacinia ante ac libero lobortis imperdiet. *Nullam mollis convallis ipsum, ac accumsan nunc vehicula vitae.* Nulla eget justo in felis tristique fringilla. Morbi sit amet tortor quis risus auctor condimentum. Morbi in ullamcorper elit. Nulla iaculis tellus sit amet mauris tempus fringilla.

Maecenas mauris lectus, lobortis et purus mattis, blandit dictum tellus.

- **Maecenas non lorem quis tellus placerat varius.** 
- *Nulla facilisi.* 
- Aenean congue fringilla justo ut aliquam. 
- [Mauris id ex erat. ](https://products.office.com/en-us/word)Nunc vulputate neque vitae justo facilisis, non condimentum ante sagittis. 
- Morbi viverra semper lorem nec molestie. 
- Maecenas tincidunt est efficitur ligula euismod, sit amet ornare est vulputate.

![](content.002.png)








In non mauris justo. Duis vehicula mi vel mi pretium, a viverra erat efficitur. Cras aliquam est ac eros varius, id iaculis dui auctor. Duis pretium neque ligula, et pulvinar mi placerat et. Nulla nec nunc sit amet nunc posuere vestibulum. Ut id neque eget tortor mattis tristique. Donec ante est, blandit sit amet tristique vel, lacinia pulvinar arcu. Pellentesque scelerisque fermentum erat, id posuere justo pulvinar ut. Cras id eros sed enim aliquam lobortis. Sed lobortis nisl ut eros efficitur tincidunt. Cras justo mi, porttitor quis mattis vel, ultricies ut purus. Ut facilisis et lacus eu cursus.

In eleifend velit vitae libero sollicitudin euismod. Fusce vitae vestibulum velit. Pellentesque vulputate lectus quis pellentesque commodo. Aliquam erat volutpat. Vestibulum in egestas velit. Pellentesque fermentum nisl vitae fringilla venenatis. Etiam id mauris vitae orci maximus ultricies. 

# **Cras fringilla ipsum magna, in fringilla dui commodo a.**


||Lorem ipsum|Lorem ipsum|Lorem ipsum|
| :- | :- | :- | :- |
|1|In eleifend velit vitae libero sollicitudin euismod.|Lorem||
|2|Cras fringilla ipsum magna, in fringilla dui commodo a.|Ipsum||
|3|Aliquam erat volutpat. |Lorem||
|4|Fusce vitae vestibulum velit. |Lorem||
|5|Etiam vehicula luctus fermentum.|Ipsum||

Etiam vehicula luctus fermentum. In vel metus congue, pulvinar lectus vel, fermentum dui. Maecenas ante orci, egestas ut aliquet sit amet, sagittis a magna. Aliquam ante quam, pellentesque ut dignissim quis, laoreet eget est. Aliquam erat volutpat. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut ullamcorper justo sapien, in cursus libero viverra eget. Vivamus auctor imperdiet urna, at pulvinar leo posuere laoreet. Suspendisse neque nisl, fringilla at iaculis scelerisque, ornare vel dolor. Ut et pulvinar nunc. Pellentesque fringilla mollis efficitur. Nullam venenatis commodo imperdiet. Morbi velit neque, semper quis lorem quis, efficitur dignissim ipsum. Ut ac lorem sed turpis imperdiet eleifend sit amet id sapien.


# **Lorem ipsum dolor sit amet, consectetur adipiscing elit.** 

Nunc ac faucibus odio. Vestibulum neque massa, scelerisque sit amet ligula eu, congue molestie mi. Praesent ut varius sem. Nullam at porttitor arcu, nec lacinia nisi. Ut ac dolor vitae odio interdum condimentum. Vivamus dapibus sodales ex, vitae malesuada ipsum cursus convallis. Maecenas sed egestas nulla, ac condimentum orci. Mauris diam felis, vulputate ac suscipit et, iaculis non est. Curabitur semper arcu ac ligula semper, nec luctus nisl blandit. Integer lacinia ante ac libero lobortis imperdiet. Nullam mollis convallis ipsum, ac accumsan nunc vehicula vitae. Nulla eget justo in felis tristique fringilla. Morbi sit amet tortor quis risus auctor condimentum. Morbi in ullamcorper elit. Nulla iaculis tellus sit amet mauris tempus fringilla.
## **Maecenas mauris lectus, lobortis et purus mattis, blandit dictum tellus.** 
Maecenas non lorem quis tellus placerat varius. Nulla facilisi. Aenean congue fringilla justo ut aliquam. Mauris id ex erat. Nunc vulputate neque vitae justo facilisis, non condimentum ante sagittis. Morbi viverra semper lorem nec molestie. Maecenas tincidunt est efficitur ligula euismod, sit amet ornare est vulputate.

In non mauris justo. Duis vehicula mi vel mi pretium, a viverra erat efficitur. Cras aliquam est ac eros varius, id iaculis dui auctor. Duis pretium neque ligula, et pulvinar mi placerat et. Nulla nec nunc sit amet nunc posuere vestibulum. Ut id neque eget tortor mattis tristique. Donec ante est, blandit sit amet tristique vel, lacinia pulvinar arcu. Pellentesque scelerisque fermentum erat, id posuere justo pulvinar ut. Cras id eros sed enim aliquam lobortis. Sed lobortis nisl ut eros efficitur tincidunt. Cras justo mi, porttitor quis mattis vel, ultricies ut purus. Ut facilisis et lacus eu cursus.
## **In eleifend velit vitae libero sollicitudin euismod.** 
Fusce vitae vestibulum velit. Pellentesque vulputate lectus quis pellentesque commodo. Aliquam erat volutpat. Vestibulum in egestas velit. Pellentesque fermentum nisl vitae fringilla venenatis. Etiam id mauris vitae orci maximus ultricies. Cras fringilla ipsum magna, in fringilla dui commodo a.

Etiam vehicula luctus fermentum. In vel metus congue, pulvinar lectus vel, fermentum dui. Maecenas ante orci, egestas ut aliquet sit amet, sagittis a magna. Aliquam ante quam, pellentesque ut dignissim quis, laoreet eget est. Aliquam erat volutpat. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut ullamcorper justo sapien, in cursus libero viverra eget. Vivamus auctor imperdiet urna, at pulvinar leo posuere laoreet. Suspendisse neque nisl, fringilla at iaculis scelerisque, ornare vel dolor. Ut et pulvinar nunc. Pellentesque fringilla mollis efficitur. Nullam venenatis commodo imperdiet. Morbi velit neque, semper quis lorem quis, efficitur dignissim ipsum. Ut ac lorem sed turpis imperdiet eleifend sit amet id sapien.

![](content.003.jpeg)
**Created with an evaluation copy of Aspose.Words. To discover the full versions of our APIs please visit: https://products.aspose.com/words/**

}
```
<sup><a href='/src/Tests/Samples.VerifyWord.verified.txt#L1-L70' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyWord.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyWord#00.verified.png](/src/Tests/Samples.VerifyWord%2300.verified.png):

<img src="/src/Tests/Samples.VerifyWord%2300.verified.png" width="200px">


### PowerPoint


#### Verify a file

<!-- snippet: VerifyPowerPoint -->
<a id='snippet-verifypowerpoint'></a>
```cs
[Test]
public Task VerifyPowerPoint() =>
    VerifyFile("sample.pptx");
```
<sup><a href='/src/Tests/Samples.cs#L43-L49' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypowerpoint' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Verify a Stream

<!-- snippet: VerifyPowerPointStream -->
<a id='snippet-verifypowerpointstream'></a>
```cs
[Test]
public Task VerifyPowerPointStream()
{
    var stream = new MemoryStream(File.ReadAllBytes("sample.pptx"));
    return Verify(stream, "pptx");
}
```
<sup><a href='/src/Tests/Samples.cs#L51-L60' title='Snippet source file'>snippet source</a> | <a href='#snippet-verifypowerpointstream' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->


#### Result

<!-- snippet: Samples.VerifyPowerPoint.verified.txt -->
<a id='snippet-Samples.VerifyPowerPoint.verified.txt'></a>
```txt
{
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
<sup><a href='/src/Tests/Samples.VerifyPowerPoint.verified.txt#L1-L22' title='Snippet source file'>snippet source</a> | <a href='#snippet-Samples.VerifyPowerPoint.verified.txt' title='Start of snippet'>anchor</a></sup>
<!-- endSnippet -->

[Samples.VerifyPowerPoint%2300.verified.png](/src/Tests/Samples.VerifyPowerPoint%2300.verified.png):


## File Samples

http://file-examples.com/


## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/).
