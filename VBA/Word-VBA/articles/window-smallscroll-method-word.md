---
title: Window.SmallScroll Method (Word)
keywords: vbawd10.chm157417576
f1_keywords:
- vbawd10.chm157417576
ms.prod: word
api_name:
- Word.Window.SmallScroll
ms.assetid: 36801d40-0a24-69aa-7641-40f90132556b
ms.date: 06/08/2017
---


# Window.SmallScroll Method (Word)

Scrolls a window or pane by the specified number of lines.


## Syntax

 _expression_ . **SmallScroll**( **_Down_** , **_Up_** , **_ToRight_** , **_ToLeft_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of lines to scroll the window down. A "line" corresponds to the distance scrolled by clicking the down scroll arrow on the vertical scroll bar once.|
| _Up_|Optional| **Variant**|The number of lines to scroll the window up. A "line" corresponds to the distance scrolled by clicking the up scroll arrow on the vertical scroll bar once.|
| _ToRight_|Optional| **Variant**|The number of lines to scroll the window to the right. A "line" corresponds to the distance scrolled by clicking the right scroll arrow on the horizontal scroll bar once.|
| _ToLeft_|Optional| **Variant**|The number of lines to scroll the window to the left. A "line" corresponds to the distance scrolled by clicking the left scroll arrow on the horizontal scroll bar once.|

## Remarks

This method is equivalent to clicking the scroll arrows on the horizontal and vertical scroll bars.

If Down and Up are both specified, the window is scrolled by the difference of the arguments. For example, if Down is 3 and Up is 6, the window is scrolled up three lines. Similarly, if ToLeft and ToRight are both specified, the window is scrolled by the difference of the arguments.

Any of these arguments can be a negative number. If no arguments are specified, the window is scrolled down by one line.


## Example

This example scrolls the active window down one line.


```vb
ActiveDocument.ActiveWindow.SmallScroll Down:=1
```

This example splits the active window and then scrolls up and over to the left.




```vb
With ActiveDocument.ActiveWindow 
 .Split = True 
 .SmallScroll Up:=5, ToLeft:=5 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

