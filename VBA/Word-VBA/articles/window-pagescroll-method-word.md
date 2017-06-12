---
title: Window.PageScroll Method (Word)
keywords: vbawd10.chm157417580
f1_keywords:
- vbawd10.chm157417580
ms.prod: word
api_name:
- Word.Window.PageScroll
ms.assetid: e3287c43-b759-e72f-5dd5-ec7f1af2bebc
ms.date: 06/08/2017
---


# Window.PageScroll Method (Word)

Scrolls through the specified pane or window page by page.


## Syntax

 _expression_ . **PageScroll**( **_Down_** , **_Up_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of pages to be scrolled down. If this argument is omitted, this value is assumed to be 1.|
| _Up_|Optional| **Variant**|The number of pages to be scrolled up.|

## Remarks

The  **PageScroll** method is available only if you are in print layout view or Web layout view. This method does not affect the position of the insertion point.

If Down and Up are both specified, the window is scrolled by the difference of the arguments. For example, if Down is 2 and Up is 4, the window is scrolled up two pages.


## Example

This example scrolls down three pages in the active window.


```vb
ActiveDocument.ActiveWindow.View.Type = wdPrintView 
ActiveDocument.ActiveWindow.PageScroll Down:=3
```

This example scrolls down one page in the active window.




```vb
ActiveDocument.ActiveWindow.View.Type = wdPrintView 
ActiveDocument.ActiveWindow.PageScroll
```


## See also


#### Concepts


[Window Object](window-object-word.md)

