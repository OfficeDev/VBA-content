---
title: Pane.PageScroll Method (Word)
keywords: vbawd10.chm157286505
f1_keywords:
- vbawd10.chm157286505
ms.prod: word
api_name:
- Word.Pane.PageScroll
ms.assetid: d96a9e10-6d57-14d7-5f4c-ab2aeffed22b
ms.date: 06/08/2017
---


# Pane.PageScroll Method (Word)

Scrolls through the specified window page by page.


## Syntax

 _expression_ . **PageScroll**( **_Down_** , **_Up_** )

 _expression_ Required. A variable that represents a **[Pane](pane-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of pages to be scrolled down. If this argument is omitted, this value is assumed to be 1.|
| _Up_|Optional| **Variant**|The number of pages to be scrolled up.|

## Remarks

The  **PageScroll** method is available only if you are in print layout view or Web layout view. This method does not affect the position of the insertion point.

If Down and Up are both specified, the window is scrolled by the difference of the arguments. For example, if Down is 2 and Up is 4, the window is scrolled up two pages.


## Example

This example scrolls up one page in the active pane.


```vb
ActiveDocument.ActiveWindow.View.Type = wdPrintView 
ActiveDocument.ActiveWindow.ActivePane.PageScroll Up:=1
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

