---
title: Page.OriginalPage Property (Visio)
keywords: vis_sdr.chm10951695
f1_keywords:
- vis_sdr.chm10951695
ms.prod: visio
api_name:
- Visio.Page.OriginalPage
ms.assetid: 4c4ca104-755a-8092-51e9-b78a6e45c95b
ms.date: 06/08/2017
---


# Page.OriginalPage Property (Visio)

 Returns the **Page** object that represents the original Microsoft Visio drawing page that was marked up on separate markup overlays by reviewers of the drawing. Read-only.


## Syntax

 _expression_ . **OriginalPage**

 _expression_ A variable that represents a **Page** object.


### Return Value

Page


## Remarks

If the  **Page** parent object is not a markup overlay, **OriginalPage** returns an error. To determine if a page is a markup overlay, check to see whether **Page.Type** = **visTypeMarkup** (3).

When a user clicks  **Track Markup**, Visio creates a new page of type  **visTypeMarkup** . The original page has type **visTypeForeground** or **visTypeBackground** . Each markup overlay is associated with a unique original drawing page.


## Example

This Microsoft Visual Basic for Applications (VBA) macro uses the  **OriginalPage** property to get the name of the original page that was marked up on a markup overlay and display it in the Immediate window. Before running this macro, make sure that a drawing page is displayed in the active window.


```vb
Public Sub OriginalPage_Example() 
 
 'Turn on Track Markup to make a markup overlay the active page. 
 Application.ActiveDocument.DocumentSheet.CellsSRC(visSectionObject, visRowDoc, visDocAddMarkup).FormulaU = True 
 
 'Display the name of the original page that currently is being marked up. 
 Debug.Print ActivePage.OriginalPage.Name 
 
End Sub
```


