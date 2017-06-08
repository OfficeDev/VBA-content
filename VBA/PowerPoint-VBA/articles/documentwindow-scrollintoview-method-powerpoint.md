---
title: DocumentWindow.ScrollIntoView Method (PowerPoint)
keywords: vbapp10.chm511029
f1_keywords:
- vbapp10.chm511029
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.ScrollIntoView
ms.assetid: 1eee6b36-9f01-5204-dd75-1172f2e00577
ms.date: 06/08/2017
---


# DocumentWindow.ScrollIntoView Method (PowerPoint)

Scrolls the document window so that items within a specified rectangular area are displayed in the document window or pane.


## Syntax

 _expression_. **ScrollIntoView**( **_Left_**, **_Top_**, **_Width_**, **_Height_**, **_Start_** )

 _expression_ A variable that represents a **DocumentWindow** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Long**|The horizontal distance (in points) from the left edge of the document window to the rectangle.|
| _Top_|Required|**Long**|The vertical distance (in points) from the upper part of the document window to the rectangle.|
| _Width_|Required|**Long**|The width of the rectangle (in points).|
| _Height_|Required|**Long**|The height of the rectangle (in points).|
| _Start_|Optional|**MsoTriState**|Determines the starting position of the rectangle in relation to the document window.|

## Remarks

If the bounding rectangle is larger than the document window, the  _Start_ parameter specifies which end of the rectangle displays or gets initial focus. This method cannot be used with outline or slide sorter views.

The  _Start_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The bottom right of the rectangle is to appear at the bottom right of the document window.|
|**msoTrue**|The default. The upper left of the rectangle is to appear at the upper left of the document window.|

## Example

This example brings into view a 100x200 point area beginning 50 points from the left edge of the slide, and 20 points from the upper part of the slide. The upper left corner of the rectangle is positioned at the upper left corner of the active document window.


```vb
ActiveWindow.ScrollIntoView Left:=50, Top:=20, _
    Width:=100, Height:=200
```


## See also


#### Concepts


[DocumentWindow Object](documentwindow-object-powerpoint.md)


