---
title: DocumentWindow.SmallScroll Method (PowerPoint)
keywords: vbapp10.chm511018
f1_keywords:
- vbapp10.chm511018
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.SmallScroll
ms.assetid: f6710bca-ad85-9257-061a-dbe5829d8b7b
ms.date: 06/08/2017
---


# DocumentWindow.SmallScroll Method (PowerPoint)

Scrolls through the specified document window by lines and columns.


## Syntax

 _expression_. **SmallScroll**( **_Down_**, **_Up_**, **_ToRight_**, **_ToLeft_** )

 _expression_ A variable that represents a **DocumentWindow** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional|**Long**|Specifies the number of lines to scroll down.|
| _Up_|Optional|**Long**|Specifies the number of lines to scroll up.|
| _ToRight_|Optional|**Long**|Specifies the number of columns to scroll right.|
| _ToLeft_|Optional|**Long**|Specifies the number of columns to scroll left.|

## Remarks

If no arguments are specified, this method scrolls down one line. If Down and Up are both specified, their effects are combined. For example, if Down is 2 and Up is 4, this method scrolls up two lines. Similarly, if ToRight and ToLeft are both specified, their effects are combined.

Any of the arguments can be a negative number.


## Example

This example scrolls down three lines in the active window.


```vb
Application.ActiveWindow.SmallScroll Down:=3
```


## See also


#### Concepts



[DocumentWindow Object](documentwindow-object-powerpoint.md)

