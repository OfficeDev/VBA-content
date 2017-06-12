---
title: Zoom.PageColumns Property (Word)
keywords: vbawd10.chm161873923
f1_keywords:
- vbawd10.chm161873923
ms.prod: word
api_name:
- Word.Zoom.PageColumns
ms.assetid: b515af7b-c579-97aa-8278-8b2ad96f8602
ms.date: 06/08/2017
---


# Zoom.PageColumns Property (Word)

Returns or sets the number of pages to be displayed side by side on-screen at the same time in print layout view or print preview. Read/write  **Long** .


## Syntax

 _expression_ . **PageColumns**

 _expression_ An expression that returns a **[Zoom](zoom-object-word.md)** object.


## Example

This example switches the active window to print layout view and displays two pages side by side.


```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .Zoom.PageColumns = 2 
 .Zoom.PageRows = 1 
End With
```

This example switches the document window for Hello.doc to print layout view and displays one full page.




```vb
With Windows("Hello.doc").View 
 .Type = wdPrintView 
 With .Zoom 
 .PageColumns = 1 
 .PageRows = 1 
 .PageFit = wdPageFitFullPage 
 End With 
End With
```


## See also


#### Concepts


[Zoom Object](zoom-object-word.md)

