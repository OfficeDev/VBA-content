---
title: Slides.InsertFromFile Method (PowerPoint)
keywords: vbapp10.chm530006
f1_keywords:
- vbapp10.chm530006
ms.prod: powerpoint
api_name:
- PowerPoint.Slides.InsertFromFile
ms.assetid: b8c6faa4-b77a-1237-cb90-00a2814e6aaa
ms.date: 06/08/2017
---


# Slides.InsertFromFile Method (PowerPoint)

Inserts slides from a file into a presentation, at the specified location. Returns an  **Integer** that represents the number of slides inserted.


## Syntax

 _expression_. **InsertFromFile**( **_FileName_**, **_Index_**, **_SlideStart_**, **_SlideEnd_** )

 _expression_ A variable that represents a **Slides** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file that contains the slides you want to insert.|
| _Index_|Required|**Long**|The index number of the  **Slide** object in the specified **Slides** collection you want to insert the new slides after.|
| _SlideStart_|Optional|**Long**|The index number of the first  **Slide** object in the **Slides** collection in the file denoted by FileName.|
| _SlideEnd_|Optional|**Long**|The index number of the last  **Slide** object in the **Slides** collection in the file denoted by FileName.|

### Return Value

Integer


## Example

This example inserts slides three through six from C:\Ppt\Sales.ppt after slide two in the active presentation.


```vb
ActivePresentation.Slides.InsertFromFile _
    "c:\ppt\sales.ppt", 2, 3, 6
```


## See also


#### Concepts


[Slides Object](slides-object-powerpoint.md)

