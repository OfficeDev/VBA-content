---
title: Panes Object (PowerPoint)
keywords: vbapp10.chm630000
f1_keywords:
- vbapp10.chm630000
ms.prod: powerpoint
api_name:
- PowerPoint.Panes
ms.assetid: a6fe4d77-dff2-6e90-1df6-eb281bc46fa6
ms.date: 06/08/2017
---


# Panes Object (PowerPoint)

A collection of  **[Pane](pane-object-powerpoint.md)** objects that represent the slide, outline, and notes panes in the document window for normal view, or the single pane of any other view in the document window.


## Remarks

In normal view, the  **Panes** collection contains three members. All other document window views have only a single pane, resulting in a **Panes** collection with one member.


## Example

Use the  **Panes** property to return the **Panes** collection. The following example tests for the number of panes in the active window. If the value is one, indicating any view other that normal view, then normal view is activated and the vertical pane divider is set to divide the document window at 15% outline pane and 85% slide pane.


```vb
With ActiveWindow

    If .Panes.Count = 1 Then

        .ViewType = ppViewNormal

        .SplitHorizontal = 15

    End If

End With
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

