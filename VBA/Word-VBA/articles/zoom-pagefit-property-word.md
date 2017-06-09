---
title: Zoom.PageFit Property (Word)
keywords: vbawd10.chm161873921
f1_keywords:
- vbawd10.chm161873921
ms.prod: word
api_name:
- Word.Zoom.PageFit
ms.assetid: ead399ec-f05f-0f28-4337-726fa3b04146
ms.date: 06/08/2017
---


# Zoom.PageFit Property (Word)

Returns or sets the view magnification of a window so that either the entire page is visible or the entire width of the page is visible. Read/write  **WdPageFit** .


## Syntax

 _expression_ . **PageFit**

 _expression_ Required. A variable that represents a **[Zoom](zoom-object-word.md)** object.


## Remarks

The  **wdPageFitFullPage** constant has no effect if the document isn't in print view.

When the  **PageFit** property is set to **wdPageFitBestFit** , the zoom percentage is automatically recalculated every time the document window size is changed. Setting this property to **wdPageFitNone** keeps the zoom percentage from being recalculated whenever this happens.


## Example

This example changes the magnification percentage of the window for Letter.doc so that the entire width of the text is visible.


```vb
With Windows("Letter.doc").View 
 .Type = wdNormalView 
 .Zoom.PageFit = wdPageFitBestFit 
End With
```

This example switches the active window to print view and changes the magnification so that the entire page is visible.




```vb
With ActiveDocument.ActiveWindow.View 
 .Type = wdPrintView 
 .Zoom.PageFit = wdPageFitFullPage 
End With
```


## See also


#### Concepts


[Zoom Object](zoom-object-word.md)

