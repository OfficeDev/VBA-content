---
title: Window.LargeScroll Method (Word)
keywords: vbawd10.chm157417575
f1_keywords:
- vbawd10.chm157417575
ms.prod: word
api_name:
- Word.Window.LargeScroll
ms.assetid: 380be0f2-dccf-7382-8fde-631ace1c5fee
ms.date: 06/08/2017
---


# Window.LargeScroll Method (Word)

Scrolls a window or pane by the specified number of screens.


## Syntax

 _expression_ . **LargeScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of screens to scroll the window down.|
| _Up_|Optional| **Variant**|The number of screens to scroll the window up.|
| _ToRight_|Optional| **Variant**|The number of screens to scroll the window to the right.|
| _ToLeft_|Optional| **Variant**|The number of screens to scroll the window to the left.|

## Remarks

This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.

If Down and Up are both specified, the window is scrolled by the difference of the arguments. For example, if Down is 2 and Up is 4, the window is scrolled up two screens. Similarly, if ToLeft and ToRight are both specified, the window is scrolled by the difference of the arguments.

Any of these arguments can be a negative number. If no arguments are specified, the window is scrolled down one screen.


## Example

This example scrolls the active window down one screen.


```vb
ActiveDocument.ActiveWindow.LargeScroll Down:=1
```

This example splits the active window and then scrolls up two screens and to the right one screen.




```vb
With ActiveDocument.ActiveWindow 
 .Split = True 
 .LargeScroll Up:=2, ToRight:=1 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

