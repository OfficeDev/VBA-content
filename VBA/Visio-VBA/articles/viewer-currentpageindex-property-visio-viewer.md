---
title: Viewer.CurrentPageIndex Property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.CurrentPageIndex
ms.assetid: 2a7950cf-c079-da63-676d-cf6a7e8a3600
ms.date: 06/08/2017
---


# Viewer.CurrentPageIndex Property (Visio Viewer)

Gets or sets the index of the page displayed when a drawing opens in Microsoft Visio Viewer. Read/write.


## Syntax

 _expression_. **CurrentPageIndex**

 _expression_An expression that returns a  **Viewer** object.


### Return Value

 **Long**


## Remarks

Set the  **CurrentPageIndex** value to the index of the page in the Visio drawing that you want to display. For example, to display Page-1, set the value to 1. If you do not specify a page index or if you set the value to 0, Visio Viewer displays the same page that was displayed the last time you saved the drawing.

If no drawing is loaded in Visio Viewer, the  **CurrentPageIndex** value is 0.

If the  **DocumentLoaded** property value is **True**, the  **CurrentPageIndex** value must be between 1 and the value of the **PageCount** property; if you set the property to a value that represents a nonexistent page, Visio Viewer ignores the setting.


## Example

The following code gets the index of the page displayed in Visio Viewer when a drawing opens.


```vb
 Debug.Print vsoViewer.CurrentPageIndex
```


