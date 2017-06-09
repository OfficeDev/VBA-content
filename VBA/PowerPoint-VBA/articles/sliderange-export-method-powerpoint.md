---
title: SlideRange.Export Method (PowerPoint)
keywords: vbapp10.chm532025
f1_keywords:
- vbapp10.chm532025
ms.prod: powerpoint
api_name:
- PowerPoint.SlideRange.Export
ms.assetid: a14b5d03-e6c4-486e-a97b-1c9bd1a18769
ms.date: 06/08/2017
---


# SlideRange.Export Method (PowerPoint)

Exports a range of slides, using the specified graphics filter, and saves the exported file under the specified file name.


## Syntax

 _expression_. **Export**( **_FileName_**, **_FilterName_**, **_ScaleWidth_**, **_ScaleHeight_** )

 _expression_ A variable that represents a **SlideRange** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to be exported and saved to disk. You can include a full path; if you don't, Microsoft PowerPoint creates a file in the current folder.|
| _FilterName_|Required|**String**|The graphics format in which you want to export slides. The specified graphics format must have an export filter registered in the Windows registry. You can specify either the registered extension or the registered filter name. Microsoft PowerPoint will first search for a matching extension in the registry. If no extension that matches the specified string is found, PowerPoint will look for a filter name that matches.|
| _ScaleWidth_|Optional|**Long**|The width in pixels of an exported slide.|
| _ScaleHeight_|Optional|**Long**|The height in pixels of an exported slide.|

## Remarks

Exporting a presentation doesn't set the  **[Saved](presentation-saved-property-powerpoint.md)** property of a presentation to **True**.

PowerPoint uses the specified graphics filter to save each individual slide. The names of the slides exported and saved to disk are determined by PowerPoint. They are typically saved by using the following naming convention: Slide1.wmf, Slide2.wmf. The path of the saved files is specified in the FileName argument.


## See also


#### Concepts


[SlideRange Object](sliderange-object-powerpoint.md)

