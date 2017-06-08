---
title: Window.UsableHeight Property (Word)
keywords: vbawd10.chm157417504
f1_keywords:
- vbawd10.chm157417504
ms.prod: word
api_name:
- Word.Window.UsableHeight
ms.assetid: 7b6458ba-41fa-d742-74e7-a606eb862a70
ms.date: 06/08/2017
---


# Window.UsableHeight Property (Word)

Returns the height (in points) of the active working area in the specified document window. Read-only  **Long** . .


## Syntax

 _expression_ . **UsableHeight**

 _expression_ A variable that represents a **[Window](window-object-word.md)** object.


## Remarks

If none of the working area is visible in the document window,  **UsableHeight** returns 1. To determine the actual available height, subtract 1 from the **UsableHeight** value.


## Example

This example displays the size of the working area in the active document window.


```vb
With ActiveDocument.ActiveWindow 
 MsgBox "Working area height = " _ 
 &; .UsableHeight &; vbLf _ 
 &; "Working area width = " _ 
 &; .UsableWidth 
End With
```


## See also


#### Concepts


[Window Object](window-object-word.md)

