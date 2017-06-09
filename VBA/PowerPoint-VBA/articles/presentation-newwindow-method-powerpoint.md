---
title: Presentation.NewWindow Method (PowerPoint)
keywords: vbapp10.chm583029
f1_keywords:
- vbapp10.chm583029
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.NewWindow
ms.assetid: 2c4e4d63-ccef-ae98-0676-fa231dec1e8c
ms.date: 06/08/2017
---


# Presentation.NewWindow Method (PowerPoint)

 Opens a new window that contains the specified presentation. Returns a **[DocumentWindow](documentwindow-object-powerpoint.md)** object that represents the new window.


## Syntax

 _expression_. **NewWindow**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

DocumentWindow


## Example

This example creates a new window that contains the contents of the active window (thereby activating the new window) and then switches back to the first window.


```vb
Set oldW = Application.ActiveWindow

Set newW = oldW.NewWindow

oldW.Activate
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

