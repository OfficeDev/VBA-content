---
title: Window.UsableWidth Property (Word)
keywords: vbawd10.chm157417503
f1_keywords:
- vbawd10.chm157417503
ms.prod: word
api_name:
- Word.Window.UsableWidth
ms.assetid: 48e8ef1a-2af2-2a3e-b879-861d6bd73af3
ms.date: 06/08/2017
---


# Window.UsableWidth Property (Word)

Returns the width (in points) of the active working area in the specified document window. Read-only  **Long** .


## Syntax

 _expression_ . **UsableWidth**

 _expression_ An expression that returns a **Window** object.


## Remarks

If none of the working area is visible in the document window,  **UsableWidth** returns 1. To determine the actual available height, subtract 1 from the **UsableWidth** value.


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

