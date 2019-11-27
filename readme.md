<!--
GENERATED FILE - DO NOT EDIT
This file was generated by [MarkdownSnippets](https://github.com/SimonCropp/MarkdownSnippets).
Source File: /readme.source.md
To change this file edit the source file and then run MarkdownSnippets.
-->

# <img src="/src/icon.png" height="30px"> Verify.Aspose

[![Build status](https://ci.appveyor.com/api/projects/status/7k8hh0guut2ioak2?svg=true)](https://ci.appveyor.com/project/SimonCropp/Verify-Aspose) [![NuGet Status](https://img.shields.io/nuget/v/Verify.Aspose.svg?cacheSeconds=86400)](https://www.nuget.org/packages/Verify.Aspose/)

Extends [Verify](https://github.com/SimonCropp/Verify) to allow verification of documents via [Aspose](https://www.aspose.com/).

Converts documents (pdf, docx, xslx, and pptx) to png for verification.

An [Aspose License](https://purchase.aspose.com/policies/license-types) is required to use this tool.

<!-- toc -->
## Contents

  * [Usage](#usage)
    * [PDF](#pdf)
    * [Excel](#excel)
    * [Word](#word)
    * [PowerPoint](#powerpoint)
  * [File Samples](#file-samples)
<!-- endtoc -->



## Usage


### PDF

<!-- snippet: VerifyPdf -->
<a id='snippet-verifypdf'/></a>
```cs
[Fact]
public async Task VerifyPdf()
{
    await this.VerifyPdf("sample.pdf");
}
```
<sup>[snippet source](/src/Tests/Samples.cs#L9-L15) / [anchor](#snippet-verifypdf)</sup>
<!-- endsnippet -->

Result: [Samples.VerifyPdf_01.verified.png](/src/Tests/Samples.VerifyPdf_01.verified.png):

<img src="/src/Tests/Samples.VerifyPdf_01.verified.png" width="200px">


### Excel

<!-- snippet: VerifyExcel -->
<a id='snippet-verifyexcel'/></a>
```cs
[Fact]
public async Task VerifyExcel()
{
    await this.VerifyExcel("sample.xlsx");
}
```
<sup>[snippet source](/src/Tests/Samples.cs#L27-L33) / [anchor](#snippet-verifyexcel)</sup>
<!-- endsnippet -->

Result: [Samples.VerifyExcel_01.01.verified.png](/src/Tests/Samples.VerifyExcel_01.01.verified.png):

<img src="/src/Tests/Samples.VerifyExcel_01.01.verified.png" width="200px">


### Word

<!-- snippet: VerifyWord -->
<a id='snippet-verifyword'/></a>
```cs
[Fact]
public async Task VerifyWord()
{
    await this.VerifyWord("sample.docx");
}
```
<sup>[snippet source](/src/Tests/Samples.cs#L35-L41) / [anchor](#snippet-verifyword)</sup>
<!-- endsnippet -->

Result: [Samples.VerifyWord_01.verified.png](/src/Tests/Samples.VerifyWord_01.verified.png):

<img src="/src/Tests/Samples.VerifyWord_01.verified.png" width="200px">


### PowerPoint

<!-- snippet: VerifyPowerPoint -->
<a id='snippet-verifypowerpoint'/></a>
```cs
[Fact]
public async Task VerifyPowerPoint()
{
    await this.VerifyPowerPoint("sample.pptx");
}
```
<sup>[snippet source](/src/Tests/Samples.cs#L18-L24) / [anchor](#snippet-verifypowerpoint)</sup>
<!-- endsnippet -->

Result: [Samples.VerifyPowerPoint_01.verified.svg](/src/Tests/Samples.VerifyPowerPoint_01.verified.svg):


## File Samples

http://file-examples.com/


## Release Notes

See [closed milestones](../../milestones?state=closed).


## Icon

[Swirl](https://thenounproject.com/term/swirl/1568686/) designed by [creativepriyanka](https://thenounproject.com/creativepriyanka) from [The Noun Project](https://thenounproject.com/creativepriyanka).
