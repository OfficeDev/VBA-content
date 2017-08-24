---
title: View.ScrollShapeIntoView Method (Publisher)
keywords: vbapb10.chm327685
f1_keywords:
- vbapb10.chm327685
ms.prod: publisher
api_name:
- Publisher.View.ScrollShapeIntoView
ms.assetid: 1d654fd4-d3b8-49e4-731d-fed27e6e0d8d
ms.date: 06/08/2017
---


# View.ScrollShapeIntoView Method (Publisher)

Scrolls the publication window so that the specified shape is displayed in the publication window or pane.


## Syntax

 _expression_. **ScrollShapeIntoView**( **_Shape_**)

 _expression_A variable that represents a  **View** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Shape|Required| **Shape**|The shape to scroll into view.|

## Example

This example adds a shape to a new page and scrolls the current view to the new shape.


```vb
Sub ScrollIntoView() 
 Dim shpStar As Shape 
 Dim intWidth As Integer 
 Dim intHeight As Integer 
 
 With ActiveDocument 
 intWidth = .PageSetup.PageWidth 
 intWidth = (intWidth / 2) - 75 
 intHeight = .PageSetup.PageHeight 
 intHeight = (intHeight / 2) - 75 
 
 With .Pages.Add(Count:=1, After:=ActiveDocument.Pages.Count) 
 Set shpStar = .Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=intWidth, Top:=intHeight, Width:=150, Height:=150) 
 shpStar.TextFrame.TextRange.Text = "New Star Shape" 
 End With 
 End With 
 
 ActiveView.ScrollShapeIntoView Shape:=shpStar 
 
End Sub
```


