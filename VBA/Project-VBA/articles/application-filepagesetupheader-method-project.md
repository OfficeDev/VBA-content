---
title: Application.FilePageSetupHeader Method (Project)
keywords: vbapj.chm2357
f1_keywords:
- vbapj.chm2357
ms.prod: project-server
api_name:
- Project.Application.FilePageSetupHeader
ms.assetid: e41ff9e1-d656-14fe-3d81-deef3065d11d
ms.date: 06/08/2017
---


# Application.FilePageSetupHeader Method (Project)

Sets up headers for printing.


## Syntax

 _expression_. **FilePageSetupHeader**( ** _Name_**, ** _Alignment_**, ** _Text_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the view or report for which to set up headers for printing.|
| _Alignment_|Optional|**Long**|The alignment of the text in the header. Can be one of the following  **PjAlignment** constants: **pjLeft**, **pjCenter**, or **pjRight**. The default value is **pjCenter**.|
| _Text_|Optional|**String**|The text to display in the header. The following special format codes may be included as part of the header:

|**Format Code**|**Description**|
|:-----|:-----|
|&;B|Turns bold printing on or off.|
|&;I|Turns italic printing on or off.|
|&;U|Turns underline printing on or off.|
|&;""fontname""|Prints characters that follow the format code in the specified font. An example would be &;""Arial"".|
|&;nn|Prints characters that follow the format code in the specified font size. Use a two-digit number to specify a size in points. An example would be &;08.|
|&;P""path""|Inserts the specified image. An example would be &;P"" _[My Documents]_ \Image.gif"". The term _[My Documents]_ represents the full path to your My Documents folder.|
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
|&;[ _Field_Name_ ]|Prints the value of the field specified with  _Field_Name_. If a macro will be run in more than one language, the field specified with _Field_Name_ must use the name localized for each language. An example would be &;[Actual Cost].|
|

### Return Value

 **Boolean**


## Remarks

Using the  **FilePageSetupHeader** method without specifying any arguments displays the **Page Setup** dialog box with the **Header** tab selected.


## Example

The following example sets up a header for printing.


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


