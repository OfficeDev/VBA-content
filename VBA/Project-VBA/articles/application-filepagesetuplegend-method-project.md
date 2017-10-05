---
title: Application.FilePageSetupLegend Method (Project)
keywords: vbapj.chm2359
f1_keywords:
- vbapj.chm2359
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupLegend
ms.assetid: b4118a37-f777-b806-9bb4-3f7e6766eda7
ms.date: 06/08/2017
---


# Application.FilePageSetupLegend Method (Project)

Creates a legend for printing views of task data.


## Syntax

_expression_. **FilePageSetupLegend** (**_Name_**, **_TextWidth_**, **_LegendOn_**, **_Alignment_**, **_Text_**, **_LabelFontName_**, **_LabelFontSize_**, **_LabelFontBold_**, **_LabelFontItalic_**, **_LabelFontUnderline_**, **_LabelFontColor_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view or report for which to set up legends for printing.|
| _TextWidth_|Optional|**Integer**|The width of the text, in inches or centimeters.|
| _LegendOn_|Optional|**Long**|The pages on which the legend appears. Can be one of the following **[PjLegend](pjlegend-enumeration-project.md)** constants: **pjNoLegend**, **pjAfterLastPage**, or **pjOnEveryPage**.|
| _Alignment_|Optional|**Long**|The alignment of the text in the legend. Can be one of the following **[PjAlignment](pjalignment-enumeration-project.md)** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjCenter**.|
| _Text_|Optional|**String**|The text to display in the legend. The special [format codes](#format-codes) can be included as part of the legend.| 
| _LabelFontName_|Optional|**String**|The name of the font.|
| _LabelFontSize_|Optional|**Integer**|The size of the font in points.|
| _LabelFontBold_|Optional|**Boolean**|**True** if the font is bold; otherwise, **False**.|
| _LabelFontItalic_|Optional|**Boolean**|**True** if the font is italic; otherwise, **False**.|
| _LabelFontUnderline_|Optional|**Boolean**|**True** if the font is underlined; otherwise, **False**.|
| _LabelFontColor_|Optional|**Long**|The color of the font. Can be one the **[PjColor](pjcolor-enumeration-project.md)** constants; see the [PjColor constants](#pjcolor-constants) table.|

<br/>

#### Format codes

|**Format code**|**Description**|
|:-----|:-----|
|&;B|Turns bold printing on or off.|
|&;I|Turns italic printing on or off.|
|&;U|Turns underline printing on or off.|
|&;""fontname""|Prints characters that follow the format code in the specified font. For example, &;""Arial"".|
|&;nn|Prints characters that follow the format code in the specified font size. Use a two-digit number to specify a size in points. For example, &;08.|
|&;P""path""|Inserts the specified image. For example, &;P"" _[My Documents]_ \Image.gif"". The term _[My Documents]_ represents the full path to your My Documents folder.|
|&;[Date]|Prints the current system date.|
|&;[Time]|Prints the current system time.|
|&;[File]|Prints the file name.|
|&;[Page]|Prints the page number.|
|&;[Pages]|Prints the total number of pages in the document.|
|&;[Project Title]|Prints the title.|
|&;[Company]|Prints the company name.|
|&;[Manager]|Prints the manager name.|
|&;[Start Date]|Prints the project start date.|
|&;[Finish Date]|Prints the project finish date.|
|&;[Current Date]|Prints the project current date.|
|&;[Status Date]|Prints the project status date.|
|&;[View]|Prints the view name.|
|&;[Report]|Prints the report name.|
|&;[Filter]|Prints the filter name.|
|&;[Saved Date]|Prints the last saved date.|
|&;[Subject]|Prints the subject.|
|&;[Author]|Prints the author.|
|&;[Keyword]|Prints the keyword(s).|
|&;[_Field_Name_]|Prints the value of the field specified with _Field_Name_. If a macro will be run in more than one language, the field specified with _Field_Name_ must use the name localized for each language. For example, &;[Actual Cost].|

<br/>

#### PjColor constants

|||
|:-----|:-----|
|**pjColorAutomatic**|**pjNavy**|
|**pjAqua**|**pjOlive**|
|**pjBlack**|**pjPurple**|
|**pjBlue**|**pjRed**|
|**pjFuchsia**|**pjSilver**|
|**pjGray**|**pjTeal**|
|**pjGreen**|**pjYellow**|
|**pjLime**|**pjWhite**|
|**pjMaroon**||

<br/>

### Return value

 **Boolean**

## Remarks

Using the **FilePageSetupLegend** method without specifying any arguments displays the **Page Setup** dialog box with the **Legend** tab selected.

**Note**   **FilePageSetupLegend** works only with views of task data.

To edit a legend where the _LabelFontColor_ can be a hexadecimal RGB value, use the **[FilePageSetupLegendEx](application-filepagesetuplegendex-method-project.md)** method.

## Example

The following example sets up a legend for printing.

```vb
Sub SetLegend() 
 
 Dim strLegend As String 
 
 strLegend = GetFontFormatCode("Arial") 
 strLegend = strLegend &; "&;BThis text will appear in the legend.&;B" 
 
 Application.FilePageSetupLegend Text:=strLegend, _ 
 Alignment:=pjCenter, LegendOn:=pjOnEveryPage 
End Sub 
 
Public Function GetFontFormatCode(strFontName As String) As String 
 
 GetFontFormatCode = "&;" &; Chr(34) &; strFontName &; Chr(34) 
End Function
```


