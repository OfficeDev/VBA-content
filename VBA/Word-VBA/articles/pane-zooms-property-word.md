---
title: Pane.Zooms Property (Word)
keywords: vbawd10.chm157286407
f1_keywords:
- vbawd10.chm157286407
ms.prod: word
api_name:
- Word.Pane.Zooms
ms.assetid: 6a09981c-cc68-2468-f750-18cb8524767c
ms.date: 06/08/2017
---


# Pane.Zooms Property (Word)

Returns a  **[Zooms](zooms-object-word.md)** collection that represents the magnification options for each view (such as normal view, outline view or print layout view).


## Syntax

 _expression_ . **Zooms**

 _expression_ An expression that returns a **[Pane](pane-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example sets the magnification in normal view to 100 percent for each open window.


```vb
Dim wndLoop as Window 
 
For Each wndLoop In Windows 
 wndLoop.ActivePane.Zooms(wdNormalView).Percentage = 100 
Next wndLoop
```

This example sets the magnification in print layout view so that an entire page is visible.




```vb
ActiveDocument.ActiveWindow.Panes(1).Zooms(wdPrintView).PageFit = _ 
 wdPageFitFullPage
```


## See also


#### Concepts


[Pane Object](pane-object-word.md)

