---
title: PageSetup.NotesOrientation Property (PowerPoint)
keywords: vbapp10.chm527007
f1_keywords:
- vbapp10.chm527007
ms.prod: powerpoint
api_name:
- PowerPoint.PageSetup.NotesOrientation
ms.assetid: 1a8e233a-58da-1296-da1f-cf59892e518f
ms.date: 06/08/2017
---


# PageSetup.NotesOrientation Property (PowerPoint)

Returns or sets the on-screen and printed orientation of notes pages, handouts, and outlines for the specified presentation. Read/write.


## Syntax

 _expression_. **NotesOrientation**

 _expression_ A variable that represents a **PageSetup** object.


### Return Value

MsoOrientation


## Remarks

The value returned by the  **NotesOrientation** property can be one of these **MsoOrientation** constants.


||
|:-----|
|**msoOrientationHorizontal**|
|**msoOrientationMixed**|
|**msoOrientationVertical**|

## Example

This example sets the orientation of all notes pages, handouts, and outlines in the active presentation to horizontal (landscape).


```vb
Application.ActivePresentation.PageSetup.NotesOrientation = msoOrientationHorizontal
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-powerpoint.md)

