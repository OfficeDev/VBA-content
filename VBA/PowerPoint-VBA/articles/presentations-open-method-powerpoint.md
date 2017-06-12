---
title: Presentations.Open Method (PowerPoint)
keywords: vbapp10.chm522006
f1_keywords:
- vbapp10.chm522006
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.Open
ms.assetid: c19456ba-e5a8-83da-00ae-dd387e38febf
ms.date: 06/08/2017
---


# Presentations.Open Method (PowerPoint)

Opens the specified presentation. Returns a  **[Presentation](presentation-object-powerpoint.md)** object that represents the opened presentation.


## Syntax

 _expression_. **Open**( **_FileName_**, **_ReadOnly_**, **_Untitled_**, **_WithWindow_** )

 _expression_ A variable that represents an **Presentations** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to open.|
| _ReadOnly_|Optional|**MsoTriState**|Specifies whether the file is opened with read/write or read-only status.|
| _Untitled_|Optional|**MsoTriState**|Specifies whether the file has a title.|
| _WithWindow_|Optional|**MsoTriState**|Specifies whether the file is visible.|

### Return Value

Presentation


## Remarks

With the proper file converters installed, Microsoft Office PowerPoint 2003 and earlier versions open files with the following MS-DOS filename extensions: .ch3, .cht, .doc, .htm, .html, .mcw, .pot, .ppa, .pps, .ppt, .pre, .rtf, .sh3, .shw, .txt, .wk1, .wk3, .wk4, .wpd, .wpf, .wps, and .xls. PowerPoint also opens files with the following filename extensions: .docm, .docx, .mhtml, .potm, .potx, .ppam, .pptm, .pptx, .ppsm, .ppsx, .thmx, .xlsm, and .xlsx.

The ReadOnly parameter value can be one of these  **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. Opens the file with read/write status.|
|**msoTrue**| Opens the file with read-only status.|
The  _Untitled_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The default. The file name automatically becomes the title of the opened presentation.|
|**msoTrue**|Opens the file without a title. This is equivalent to creating a copy of the file.|
The  _WithWindow_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| Hides the opened presentation.|
|**msoTrue**|The default. Opens the file in a visible window.|

## Example

This example opens a presentation with read-only status.


```
Presentations.Open FileName:="C:\My Documents\pres1.pptx", ReadOnly:=msoTrue
```


## See also


#### Concepts


[Presentations Object](presentations-object-powerpoint.md)

